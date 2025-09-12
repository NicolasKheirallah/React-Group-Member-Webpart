import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IUser, ISite, IGroup } from '../types/interfaces';
import { CacheService } from './CacheService';
import { GroupMemberService } from './GroupMemberService';

// Type for Microsoft Graph Client
interface IMSGraphClient {
  api(path: string): IMSGraphClientRequest;
}

interface IMSGraphClientRequest {
  select(properties: string): IMSGraphClientRequest;
  expand(properties: string): IMSGraphClientRequest;
  filter(filter: string): IMSGraphClientRequest;
  get(): Promise<{ value?: unknown[] } & Record<string, unknown>>;
}

export interface ISitePermissionService {
  getCurrentSite(): Promise<ISite | undefined>;
  getSiteMembers(siteId: string): Promise<IUser[]>;
  getAllSiteMembers(): Promise<IUser[]>;
}

export class SitePermissionService implements ISitePermissionService {
  private context: WebPartContext;
  private graphClient: IMSGraphClient | undefined;
  private cacheService: CacheService;
  private groupMemberService: GroupMemberService;

  constructor(context: WebPartContext) {
    this.context = context;
    this.cacheService = CacheService.getInstance();
    this.groupMemberService = new GroupMemberService(context);
  }

  private async getGraphClient(): Promise<IMSGraphClient> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  public async getCurrentSite(): Promise<ISite | undefined> {
    try {
      const client = await this.getGraphClient();
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Get site information using the current site URL
      const hostname = new URL(siteUrl).hostname;
      const sitePath = new URL(siteUrl).pathname;
      
      const response = await client.api(`/sites/${hostname}:${sitePath}`).get();
      
      const r = response as Record<string, unknown>;
      const sharepointIds = r.sharepointIds as Record<string, unknown> || {};
      return {
        id: r.id as string,
        displayName: r.displayName as string,
        webUrl: r.webUrl as string,
        siteCollectionId: sharepointIds.siteId as string,
        webId: sharepointIds.webId as string
      };
    } catch (error) {
      console.warn('Could not get current site info:', error);
      return undefined;
    }
  }

  public async getSiteMembers(siteId: string): Promise<IUser[]> {
    const cacheKey = `siteMembers_${siteId}`;
    
    // Check LRU cache first
    const cachedData = this.cacheService.getUserData(cacheKey);
    if (cachedData) {
      return cachedData as IUser[];
    }

    try {
      const client = await this.getGraphClient();
      const allUsers: IUser[] = [];

      // Method 1: Get site permissions (includes inherited permissions)
      try {
        const permissionsResponse = await client
          .api(`/sites/${siteId}/permissions`)
          .select('id,roles,grantedToIdentitiesV2,grantedTo,inheritedFrom')
          .get();

        for (const permission of permissionsResponse.value || []) {
          const p = permission as Record<string, unknown>;
          const roles = p.roles as string[] || [];
          const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
          
          if (grantedTo) {
            for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
              const id = identity as Record<string, unknown>;
              const accessLevel = this.mapSiteRolesToAccessLevel(roles);
              
              if (id.user) {
                // Direct user permission
                const user = id.user as Record<string, unknown>;
                allUsers.push({
                  id: user.id as string,
                  displayName: user.displayName as string,
                  mail: user.mail as string,
                  userPrincipalName: user.userPrincipalName as string,
                  jobTitle: user.jobTitle as string,
                  department: user.department as string,
                  officeLocation: user.officeLocation as string,
                  accessLevel,
                  source: 'site'
                });
              } else if (id.group) {
                // Security group or M365 group permission - resolve members
                try {
                  const group = id.group as Record<string, unknown>;
                  const groupMembers = await this.groupMemberService.resolveGroupMembers(group.id as string, accessLevel);
                  allUsers.push(...groupMembers);
                } catch (error) {
                  console.warn(`Failed to resolve group members for group ${id.group}:`, error);
                }
              }
            }
          }
        }
      } catch (error) {
        console.warn('Failed to get site permissions via /permissions endpoint:', error);
      }

      // Method 2: Try SharePoint REST API for additional members (fallback)
      try {
        const alternativeMembers = await this.getSharePointSiteMembers(siteId);
        allUsers.push(...alternativeMembers);
      } catch (error) {
        console.warn('Failed to get members via SharePoint REST API:', error);
      }

      // Method 3: Try to get site administrators specifically
      try {
        const siteAdmins = await this.getSiteAdministrators(siteId);
        allUsers.push(...siteAdmins);
      } catch (error) {
        console.warn('Failed to get site administrators:', error);
      }

      // Deduplicate users by ID, giving priority to higher access levels
      const userMap = new Map<string, IUser>();
      
      for (const user of allUsers) {
        const existingUser = userMap.get(user.id);
        if (!existingUser || this.getAccessLevelPriority(user.accessLevel) > this.getAccessLevelPriority(existingUser.accessLevel)) {
          userMap.set(user.id, user);
        }
      }

      const uniqueUsers = Array.from(userMap.values());

      // Cache the result
      this.cacheService.setUserData(cacheKey, uniqueUsers);
      
      return uniqueUsers;
    } catch (error) {
      console.error(`Error fetching site members for ${siteId}:`, error);
      return [];
    }
  }

  public async getAllSiteMembers(): Promise<IUser[]> {
    try {
      const currentSite = await this.getCurrentSite();
      if (!currentSite) {
        console.warn('Could not determine current site');
        return [];
      }

      const allUsers: IUser[] = [];
      const errors: string[] = [];

      // COMPREHENSIVE APPROACH: Get ALL users from multiple sources
      
      // 1. Try multiple approaches to find associated M365 group
      const associatedGroup = await this.findAssociatedGroup(currentSite);

      // 2. Get M365 group members if there's an associated group
      if (associatedGroup) {
        console.log(`Found associated M365 group: ${associatedGroup.displayName} (${associatedGroup.id})`);
        try {
          // Get group owners
          try {
            const groupOwners = await this.groupMemberService.getGroupMembers(associatedGroup.id, 'admin');
            const ownersWithLevel = groupOwners.map(user => ({ 
              ...user, 
              accessLevel: 'owner' as const, 
              source: 'group' as const 
            }));
            allUsers.push(...ownersWithLevel);
            console.log(`Added ${ownersWithLevel.length} M365 group owners`);
          } catch (error) {
            errors.push(`Failed to get owners from group: ${error}`);
          }

          // Get group members
          try {
            const groupMembers = await this.groupMemberService.getGroupMembers(associatedGroup.id, 'member');
            const membersWithLevel = groupMembers.map(user => ({ 
              ...user, 
              accessLevel: 'member' as const, 
              source: 'group' as const 
            }));
            allUsers.push(...membersWithLevel);
            console.log(`Added ${membersWithLevel.length} M365 group members`);
          } catch (error) {
            errors.push(`Failed to get members from group: ${error}`);
          }
        } catch (error) {
          errors.push(`Failed to get group members: ${error}`);
        }
      } else {
        console.log('No associated M365 group found - this is likely a Communication Site');
      }

      // 3. ALWAYS get direct site permissions (SharePoint groups, individual permissions)
      console.log('Getting SharePoint site permissions...');
      try {
        const siteMembers = await this.getComprehensiveSiteMembers(currentSite.id);
        allUsers.push(...siteMembers);
        console.log(`Added ${siteMembers.length} users from SharePoint site permissions`);
      } catch (error) {
        errors.push(`Failed to get site members: ${error}`);
        console.warn('Site members retrieval failed:', error);
      }

      // 4. Try SharePoint REST API for additional discovery
      try {
        const restApiMembers = await this.getSharePointRestApiMembers();
        allUsers.push(...restApiMembers);
        console.log(`Added ${restApiMembers.length} users from SharePoint REST API`);
      } catch (error) {
        errors.push(`Failed to get SharePoint REST API members: ${error}`);
        console.warn('SharePoint REST API members retrieval failed:', error);
      }

      // 5. Get users from all SharePoint groups (this is the key missing piece!)
      try {
        const sharePointGroupMembers = await this.getAllSharePointGroupMembers(currentSite.id);
        allUsers.push(...sharePointGroupMembers);
        console.log(`Added ${sharePointGroupMembers.length} users from SharePoint groups`);
      } catch (error) {
        errors.push(`Failed to get SharePoint group members: ${error}`);
        console.warn('SharePoint group members retrieval failed:', error);
      }

      // 6. For Communication sites, get visitors and external users
      try {
        const visitors = await this.getSiteVisitors(currentSite.id);
        allUsers.push(...visitors);
        console.log(`Added ${visitors.length} site visitors`);
      } catch (error) {
        errors.push(`Failed to get site visitors: ${error}`);
      }

      // 7. Try getting all user information profiles that have access
      try {
        const userProfiles = await this.getUserProfilesWithAccess(currentSite.id);
        allUsers.push(...userProfiles);
        console.log(`Added ${userProfiles.length} users from user profiles`);
      } catch (error) {
        errors.push(`Failed to get user profiles: ${error}`);
      }

      // 8. CRITICAL: Try to get ALL site users using SharePoint's native user enumeration (like _layouts/15/user.aspx does)
      try {
        const allSiteUsers = await this.getAllSiteUsersFromSharePoint();
        allUsers.push(...allSiteUsers);
        console.log(`Added ${allSiteUsers.length} users from SharePoint native user enumeration`);
      } catch (error) {
        errors.push(`Failed to get SharePoint native users: ${error}`);
        console.warn('SharePoint native user enumeration failed:', error);
      }

      // 8.1. BACKUP: Try simple SharePoint REST API calls with minimal headers
      if (allUsers.length <= 5) { // If we have 5 or fewer users, try backup approach to get more
        try {
          const backupUsers = await this.getSimpleSharePointUsers();
          allUsers.push(...backupUsers);
          console.log(`Added ${backupUsers.length} users from simple SharePoint backup method`);
        } catch (error) {
          console.warn('Simple SharePoint backup method failed:', error);
        }
      }

      // 9. Try to get users from role assignments (this is what _layouts/15/user.aspx shows)
      try {
        const roleAssignmentUsers = await this.getUsersFromRoleAssignments();
        allUsers.push(...roleAssignmentUsers);
        console.log(`Added ${roleAssignmentUsers.length} users from role assignments`);
      } catch (error) {
        errors.push(`Failed to get role assignment users: ${error}`);
        console.warn('Role assignment users retrieval failed:', error);
      }

      // 8. Final fallback: get current user and try to enumerate from site usage
      if (allUsers.length === 0) {
        console.warn('No site members found through any method. Trying final fallback approaches...');
        
        try {
          const fallbackMembers = await this.getFallbackSiteMembers(currentSite.id);
          allUsers.push(...fallbackMembers);
          console.log(`Added ${fallbackMembers.length} users from fallback methods`);
        } catch (fallbackError) {
          errors.push(`All fallback methods failed: ${fallbackError}`);
        }
        
        // Add current user as absolute minimum
        const currentUser = this.context.pageContext.user;
        if (currentUser && allUsers.length === 0) {
          allUsers.push({
            id: currentUser.loginName || currentUser.email || 'current',
            displayName: currentUser.displayName,
            mail: currentUser.email,
            userPrincipalName: currentUser.loginName || currentUser.email,
            accessLevel: 'admin' as const,
            source: 'site' as const
          });
          console.log('Added current user as absolute minimum baseline');
        }
      }

      if (errors.length > 0) {
        console.warn('Some member retrieval methods failed:', errors.slice(0, 5));
      }

      const userMap = new Map<string, IUser>();
      const seenKeys = new Set<string>();
      
      for (const user of allUsers) {
        const possibleKeys = [
          user.id,
          user.userPrincipalName,
          user.mail,
          user.displayName
        ].filter((key): key is string => key !== undefined && key !== null && key.trim() !== '');
        
        if (possibleKeys.length === 0) {
          continue;
        }
        
        let isDuplicate = false;
        let existingUser: IUser | undefined;
        let keyToUse = '';
        
        for (const key of possibleKeys) {
          if (seenKeys.has(key) || userMap.has(key)) {
            isDuplicate = true;
            existingUser = userMap.get(key);
            keyToUse = key;
            break;
          }
        }
        
        if (!isDuplicate) {
          keyToUse = possibleKeys[0];
          userMap.set(keyToUse, user);
          possibleKeys.forEach(key => seenKeys.add(key));
        } else if (existingUser && this.getAccessLevelPriority(user.accessLevel) > this.getAccessLevelPriority(existingUser.accessLevel)) {
          userMap.set(keyToUse, user);
        }
      }

      const uniqueUsers = Array.from(userMap.values());
      console.log(`Found ${uniqueUsers.length} unique site members from ${allUsers.length} total entries`);
      console.log('Final user breakdown by access level:', {
        owners: uniqueUsers.filter(u => u.accessLevel === 'owner').length,
        admins: uniqueUsers.filter(u => u.accessLevel === 'admin').length,
        members: uniqueUsers.filter(u => u.accessLevel === 'member').length,
        visitors: uniqueUsers.filter(u => u.accessLevel === 'visitor').length
      });
      console.log('Unique users:', uniqueUsers.map(u => ({ 
        name: u.displayName, 
        accessLevel: u.accessLevel, 
        source: u.source,
        upn: u.userPrincipalName 
      })));
      
      return uniqueUsers;
    } catch (error) {
      console.error('Critical error in getAllSiteMembers:', error);
      return [];
    }
  }

  private mapSiteRolesToAccessLevel(roles: string[]): 'owner' | 'admin' | 'member' | 'visitor' {
    // Convert roles to lowercase for case-insensitive comparison
    const lowerRoles = roles.map(role => role.toLowerCase());
    
    // Site Owner or Full Control
    if (lowerRoles.some(role => 
      role.includes('owner') || 
      role.includes('fullcontrol') || 
      role === 'full control' ||
      role.includes('siteadmin') ||
      role.includes('site admin')
    )) {
      return 'owner';
    }
    
    // Site Administrator, Design, or Manage permissions
    if (lowerRoles.some(role => 
      role.includes('admin') || 
      role.includes('manage') || 
      role.includes('design') ||
      role === 'manage hierarchy' ||
      role === 'approve' ||
      role.includes('moderate') ||
      role.includes('restrict')
    )) {
      return 'admin';
    }
    
    // Contributors, Edit, Write permissions
    if (lowerRoles.some(role => 
      role.includes('edit') || 
      role.includes('contribute') || 
      role.includes('write') ||
      role === 'add and customize pages' ||
      role === 'add items' ||
      role === 'edit items' ||
      role.includes('create') ||
      role.includes('modify')
    )) {
      return 'member';
    }
    
    // Visitors, Read-only permissions (be more explicit about visitor roles)
    if (lowerRoles.some(role => 
      role.includes('read') || 
      role.includes('view') ||
      role.includes('visitor') ||
      role === 'view only' ||
      role === 'limited access' ||
      role.includes('browse')
    )) {
      return 'visitor';
    }
    
    // Default to visitor for any other permissions
    return 'visitor';
  }

  // Additional method to get SharePoint site members via alternative API
  private async getSharePointSiteMembers(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const allUsers: IUser[] = [];

    try {
      // Try to get site users via alternative Graph endpoints
      const siteDrive = await client.api(`/sites/${siteId}/drive`).get();
      
      if (siteDrive) {
        // Get users who have access to the site drive
        const drivePermissions = await client
          .api(`/sites/${siteId}/drive/root/permissions`)
          .select('id,roles,grantedToIdentitiesV2,grantedTo')
          .get();

        for (const permission of drivePermissions.value || []) {
          const p = permission as Record<string, unknown>;
          const roles = p.roles as string[] || [];
          const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
          
          if (grantedTo) {
            for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
              const id = identity as Record<string, unknown>;
              const accessLevel = this.mapSiteRolesToAccessLevel(roles);
              
              if (id.user) {
                const user = id.user as Record<string, unknown>;
                allUsers.push({
                  id: user.id as string,
                  displayName: user.displayName as string,
                  mail: user.mail as string,
                  userPrincipalName: user.userPrincipalName as string,
                  jobTitle: user.jobTitle as string,
                  department: user.department as string,
                  officeLocation: user.officeLocation as string,
                  accessLevel,
                  source: 'site'
                });
              }
            }
          }
        }
      }
    } catch (error) {
      console.debug('Alternative SharePoint API method failed:', error);
    }

    return allUsers;
  }

  // Method to specifically get site administrators
  private async getSiteAdministrators(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const admins: IUser[] = [];

    try {
      // Method 1: Try to get site information with owner details
      const siteInfo = await client
        .api(`/sites/${siteId}`)
        .select('id,displayName,createdBy,siteCollection')
        .get();

      const siteData = siteInfo as Record<string, unknown>;
      
      // Add site creator as owner if available
      if (siteData.createdBy) {
        const createdBy = siteData.createdBy as Record<string, unknown>;
        if (createdBy.user) {
          const user = createdBy.user as Record<string, unknown>;
          admins.push({
            id: user.id as string || 'creator',
            displayName: user.displayName as string || 'Site Creator',
            mail: user.email as string,
            userPrincipalName: user.userPrincipalName as string || user.email as string,
            accessLevel: 'owner',
            source: 'site'
          });
        }
      }

      // Method 2: Try to get site collection administrators
      try {
        const siteCollection = siteData.siteCollection as Record<string, unknown>;
        if (siteCollection && siteCollection.hostname) {
          // Additional logic to get tenant admins if needed
          console.debug('Site collection info available for admin discovery');
        }
      } catch (error) {
        console.debug('Could not get site collection admins:', error);
      }

    } catch (error) {
      console.debug('Could not get site administrators:', error);
    }

    return admins;
  }

  // Enhanced method to get visitors specifically
  private async getSiteVisitors(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const visitors: IUser[] = [];

    try {
      // Look for "Everyone" or "Everyone except external users" permissions
      const permissionsResponse = await client
        .api(`/sites/${siteId}/permissions`)
        .select('id,roles,grantedToIdentitiesV2,grantedTo')
        .get();

      for (const permission of permissionsResponse.value || []) {
        const p = permission as Record<string, unknown>;
        const roles = p.roles as string[] || [];
        const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
        
        // Only process if this is clearly a visitor-level permission
        if (this.mapSiteRolesToAccessLevel(roles) === 'visitor' && grantedTo) {
          for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
            const id = identity as Record<string, unknown>;
            
            if (id.user) {
              const user = id.user as Record<string, unknown>;
              visitors.push({
                id: user.id as string,
                displayName: user.displayName as string,
                mail: user.mail as string,
                userPrincipalName: user.userPrincipalName as string,
                jobTitle: user.jobTitle as string,
                department: user.department as string,
                officeLocation: user.officeLocation as string,
                accessLevel: 'visitor',
                source: 'site'
              });
            }
          }
        }
      }
    } catch (error) {
      console.debug('Could not get site visitors:', error);
    }

    return visitors;
  }

  private getAccessLevelPriority(level?: string): number {
    switch (level) {
      case 'owner': return 4;
      case 'admin': return 3;
      case 'member': return 2;
      case 'visitor': return 1;
      default: return 0;
    }
  }

  private async findAssociatedGroup(site: ISite): Promise<IGroup | undefined> {
    try {
      // Method 1: Try to get the site's group directly via Graph API
      try {
        const client = await this.getGraphClient();
        const response = await client.api(`/sites/${site.id}/drive`).get();
        const driveResponse = response as Record<string, unknown>;
        
        if (driveResponse.quota && (driveResponse.quota as Record<string, unknown>).deleted === undefined) {
          // This is likely a group-connected site, try to get the group
          const groupResponse = await client.api(`/sites/${site.id}/group`).get();
          const g = groupResponse as Record<string, unknown>;
          
          return {
            id: g.id as string,
            displayName: g.displayName as string,
            '@odata.type': '#microsoft.graph.group',
            description: g.description as string
          };
        }
      } catch {
        // Site might not have an associated group, continue with other methods
        console.log('No direct group association found, trying other methods');
      }

      // Method 2: Match by site URL patterns (more reliable than display name)
      const groups = await this.groupMemberService.getUserGroups();
      const siteUrl = site.webUrl.toLowerCase();
      
      // Try to find group by examining site URL structure
      const potentialGroup = groups.find(group => {
        const groupName = group.displayName.toLowerCase().replace(/\s+/g, '');
        const siteName = this.extractSiteNameFromUrl(siteUrl);
        
        return groupName === siteName || 
               siteName.includes(groupName) || 
               groupName.includes(siteName);
      });

      if (potentialGroup) {
        return potentialGroup;
      }

      // Method 3: Check if current user is owner of any groups that might match
      // This is useful for scenarios where the user has limited group visibility
      for (const group of groups) {
        try {
          const groupOwners = await this.groupMemberService.getGroupMembers(group.id, 'admin');
          const currentUser = this.context.pageContext.user;
          
          if (groupOwners.some(owner => 
            owner.userPrincipalName === currentUser.loginName ||
            owner.mail === currentUser.email
          )) {
            // User is owner of this group, might be the associated group
            const siteName = this.extractSiteNameFromUrl(siteUrl);
            const groupName = group.displayName.toLowerCase();
            
            if (groupName.includes(siteName) || siteName.includes(groupName)) {
              return group;
            }
          }
        } catch {
          // Continue to next group if this one fails
          continue;
        }
      }

      console.log('No associated M365 group found for this site');
      return undefined;
    } catch (error) {
      console.warn('Error finding associated group:', error);
      return undefined;
    }
  }

  private extractSiteNameFromUrl(siteUrl: string): string {
    try {
      const url = new URL(siteUrl);
      const pathParts = url.pathname.split('/').filter(part => part.length > 0);
      
      // For sites like /sites/sitename or /teams/teamname
      if (pathParts.length >= 2 && (pathParts[0] === 'sites' || pathParts[0] === 'teams')) {
        return pathParts[1].toLowerCase().replace(/[^a-z0-9]/g, '');
      }
      
      // For other patterns, try to extract meaningful name
      const lastPart = pathParts[pathParts.length - 1];
      return lastPart.toLowerCase().replace(/[^a-z0-9]/g, '');
    } catch {
      return '';
    }
  }

  // Removed getSiteMembersWithRetry as it's not used in the new comprehensive approach

  // COMPREHENSIVE method to get site members using multiple Graph API approaches
  private async getComprehensiveSiteMembers(siteId: string): Promise<IUser[]> {
    const client = await this.getGraphClient();
    const allUsers: IUser[] = [];

    try {
      // Method 1: Get site permissions with proper selection (no expand on grantedToIdentitiesV2)
      const permissionsResponse = await client
        .api(`/sites/${siteId}/permissions`)
        .select('id,roles,grantedToIdentitiesV2,grantedTo,inheritedFrom')
        .get();

      console.log(`Found ${permissionsResponse.value?.length || 0} permissions entries`);

      for (const permission of permissionsResponse.value || []) {
        const p = permission as Record<string, unknown>;
        const roles = p.roles as string[] || [];
        const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
        
        if (grantedTo) {
          for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
            const id = identity as Record<string, unknown>;
            const accessLevel = this.mapSiteRolesToAccessLevel(roles);
            
            if (id.user) {
              // Direct user permission
              const user = id.user as Record<string, unknown>;
              const upn = user.userPrincipalName as string || user.email as string;
              const email = user.mail as string || user.email as string;
              
              // Skip service accounts and users without proper UPN/email
              if (upn && email && 
                  !upn.includes('app@sharepoint') && 
                  !upn.includes('SHAREPOINT\\system') &&
                  !email.includes('noreply') &&
                  user.displayName &&
                  (user.displayName as string) !== 'System Account') {
                
                allUsers.push({
                  id: user.id as string,
                  displayName: user.displayName as string || email,
                  mail: email,
                  userPrincipalName: upn,
                  jobTitle: user.jobTitle as string,
                  department: user.department as string,
                  officeLocation: user.officeLocation as string,
                  accessLevel,
                  source: 'site'
                });
              }
            } else if (id.group) {
              // Security group or SharePoint group - resolve all members
              try {
                const group = id.group as Record<string, unknown>;
                const groupId = group.id as string;
                console.log(`Resolving members for group: ${group.displayName} (${groupId})`);
                
                // Try different approaches to get group members
                const groupMembers = await this.resolveAllGroupMembers(groupId, accessLevel);
                allUsers.push(...groupMembers);
                console.log(`Added ${groupMembers.length} members from group ${group.displayName}`);
              } catch (error) {
                console.warn(`Failed to resolve group members for group ${id.group}:`, error);
              }
            }
          }
        }
      }

      // Method 2: Try to get site drive permissions as well
      try {
        const drivePermissions = await client
          .api(`/sites/${siteId}/drive/root/permissions`)
          .select('id,roles,grantedToIdentitiesV2,grantedTo')
          .get();

        for (const permission of drivePermissions.value || []) {
          const p = permission as Record<string, unknown>;
          const roles = p.roles as string[] || [];
          const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
          
          if (grantedTo) {
            for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
              const id = identity as Record<string, unknown>;
              const accessLevel = this.mapSiteRolesToAccessLevel(roles);
              
              if (id.user) {
                const user = id.user as Record<string, unknown>;
                const upn = user.userPrincipalName as string || user.email as string;
                const email = user.mail as string || user.email as string;
                
                // Skip service accounts and users without proper UPN/email
                if (upn && email && 
                    !upn.includes('app@sharepoint') && 
                    !upn.includes('SHAREPOINT\\system') &&
                    !email.includes('noreply') &&
                    user.displayName &&
                    (user.displayName as string) !== 'System Account') {
                  
                  allUsers.push({
                    id: user.id as string,
                    displayName: user.displayName as string || email,
                    mail: email,
                    userPrincipalName: upn,
                    jobTitle: user.jobTitle as string,
                    department: user.department as string,
                    officeLocation: user.officeLocation as string,
                    accessLevel,
                    source: 'site'
                  });
                }
              }
            }
          }
        }
      } catch (error) {
        console.debug('Drive permissions not accessible or not available:', error);
      }

    } catch (error) {
      console.warn('Comprehensive site members retrieval failed:', error);
    }

    return allUsers;
  }

  // NEW METHOD: Get all SharePoint groups and their members
  private async getAllSharePointGroupMembers(siteId: string): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      // This is a crucial method - try to get SharePoint groups via different approaches
      
      // Approach 1: Try to get site groups via Graph API lists endpoint
      const client = await this.getGraphClient();
      
      try {
        // Get all lists to find potential group associations
        const listsResponse = await client
          .api(`/sites/${siteId}/lists`)
          .filter("listTemplate eq 'genericList' or listTemplate eq 'documentLibrary'")
          .select('id,displayName,list')
          .get();

        // For each list, try to get its permissions
        for (const listItem of listsResponse.value || []) {
          const list = listItem as Record<string, unknown>;
          try {
            const listPermissions = await client
              .api(`/sites/${siteId}/lists/${list.id}/permissions`)
              .select('id,roles,grantedToIdentitiesV2,grantedTo')
              .get();

            for (const permission of listPermissions.value || []) {
              const p = permission as Record<string, unknown>;
              const roles = p.roles as string[] || [];
              const grantedTo = p.grantedToIdentitiesV2 || p.grantedTo;
              
              if (grantedTo) {
                for (const identity of Array.isArray(grantedTo) ? grantedTo : [grantedTo]) {
                  const id = identity as Record<string, unknown>;
                  
                  if (id.group) {
                    const group = id.group as Record<string, unknown>;
                    const accessLevel = this.mapSiteRolesToAccessLevel(roles);
                    
                    try {
                      const groupMembers = await this.resolveAllGroupMembers(group.id as string, accessLevel);
                      allUsers.push(...groupMembers);
                    } catch (error) {
                      console.debug(`Could not resolve members for list group ${group.displayName}:`, error);
                    }
                  }
                }
              }
            }
          } catch (error) {
            console.debug(`Could not get permissions for list ${list.displayName}:`, error);
          }
        }
      } catch (error) {
        console.debug('Could not enumerate lists for group discovery:', error);
      }

      // Approach 2: Try SharePoint REST API directly
      try {
        const sharePointMembers = await this.getSharePointRestApiMembers();
        allUsers.push(...sharePointMembers);
      } catch (error) {
        console.debug('SharePoint REST API approach failed:', error);
      }

    } catch (error) {
      console.warn('Failed to get SharePoint group members:', error);
    }

    return allUsers;
  }

  // NEW METHOD: Enhanced group member resolution
  private async resolveAllGroupMembers(groupId: string, accessLevel: 'owner' | 'admin' | 'member' | 'visitor'): Promise<IUser[]> {
    const cacheKey = `allGroupMembers_${groupId}_${accessLevel}`;
    
    // Check cache first
    const cachedData = this.cacheService.getUserData(cacheKey);
    if (cachedData) {
      return cachedData as IUser[];
    }

    const allMembers: IUser[] = [];
    
    try {
      const client = await this.getGraphClient();
      
      // Try multiple approaches to get group members
      
      // Approach 1: Direct group members endpoint
      try {
        const response = await client
          .api(`/groups/${groupId}/members`)
          .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,userType,accountEnabled')
          .get();

        const members = response.value || [];
        for (const member of members) {
          const m = member as Record<string, unknown>;
          // Only include active user accounts
          if (m.userType !== 'Guest' || m.accountEnabled !== false) {
            allMembers.push({
              id: m.id as string,
              displayName: m.displayName as string,
              mail: m.mail as string,
              userPrincipalName: m.userPrincipalName as string,
              jobTitle: m.jobTitle as string,
              department: m.department as string,
              officeLocation: m.officeLocation as string,
              accessLevel,
              source: 'site'
            });
          }
        }
      } catch (error) {
        console.debug(`Direct group members failed for ${groupId}:`, error);
      }

      // Approach 2: Try to get group owners as well
      try {
        const ownersResponse = await client
          .api(`/groups/${groupId}/owners`)
          .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation')
          .get();

        const owners = ownersResponse.value || [];
        for (const owner of owners) {
          const o = owner as Record<string, unknown>;
          // Add owners with elevated access level
          const ownerAccessLevel = accessLevel === 'visitor' ? 'admin' : 'owner';
          allMembers.push({
            id: o.id as string,
            displayName: o.displayName as string,
            mail: o.mail as string,
            userPrincipalName: o.userPrincipalName as string,
            jobTitle: o.jobTitle as string,
            department: o.department as string,
            officeLocation: o.officeLocation as string,
            accessLevel: ownerAccessLevel,
            source: 'site'
          });
        }
      } catch (error) {
        console.debug(`Group owners failed for ${groupId}:`, error);
      }

      // Cache the result
      this.cacheService.setUserData(cacheKey, allMembers);
      
    } catch (error) {
      console.warn(`Failed to resolve all group members for ${groupId}:`, error);
    }

    return allMembers;
  }

  // NEW METHOD: SharePoint REST API approach for legacy compatibility
  private async getSharePointRestApiMembers(): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      // Use SPFx context to make SharePoint REST calls
      const spHttpClient = this.context.spHttpClient;
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Get site users
      const usersResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/siteusers?$select=Id,Title,Email,UserPrincipalName,LoginName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json'
          }
        }
      );
      
      if (usersResponse.ok) {
        const usersData = await usersResponse.json();
        const users = usersData.value || usersData.d?.results || [];
        
        for (const user of users) {
          // Skip system accounts and groups
          if (user.LoginName && !user.LoginName.includes('c:0(.s|true)') && !user.LoginName.includes('SHAREPOINT\\\\system')) {
            allUsers.push({
              id: user.LoginName || user.Id?.toString() || 'unknown',
              displayName: user.Title || user.Email || 'Unknown User',
              mail: user.Email || '',
              userPrincipalName: user.UserPrincipalName || user.Email || user.LoginName,
              accessLevel: 'member' as const, // Default to member, will be refined later
              source: 'site' as const
            });
          }
        }
        
        console.log(`SharePoint REST API found ${allUsers.length} site users`);
      }
      
      // Get site groups
      const groupsResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/sitegroups?$expand=Users&$select=Title,Users/Title,Users/Email,Users/UserPrincipalName,Users/LoginName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json'
          }
        }
      );
      
      if (groupsResponse.ok) {
        const groupsData = await groupsResponse.json();
        const groups = groupsData.value || groupsData.d?.results || [];
        
        for (const group of groups) {
          const groupTitle = group.Title || '';
          let accessLevel: 'owner' | 'admin' | 'member' | 'visitor' = 'visitor';
          
          // Determine access level based on SharePoint group name
          if (groupTitle.toLowerCase().includes('owner')) {
            accessLevel = 'owner';
          } else if (groupTitle.toLowerCase().includes('admin') || groupTitle.toLowerCase().includes('contribute')) {
            accessLevel = 'admin';
          } else if (groupTitle.toLowerCase().includes('member') || groupTitle.toLowerCase().includes('edit')) {
            accessLevel = 'member';
          }
          
          const users = group.Users?.results || [];
          for (const user of users) {
            if (user.LoginName && !user.LoginName.includes('c:0(.s|true)') && !user.LoginName.includes('SHAREPOINT\\\\system')) {
              allUsers.push({
                id: user.LoginName || 'unknown',
                displayName: user.Title || user.Email || 'Unknown User',
                mail: user.Email || '',
                userPrincipalName: user.UserPrincipalName || user.Email || user.LoginName,
                accessLevel,
                source: 'site' as const
              });
            }
          }
        }
        
        console.log(`SharePoint REST API found ${groups.length} site groups with members`);
      }
      
    } catch (error) {
      console.warn('SharePoint REST API member retrieval failed:', error);
    }
    
    return allUsers;
  }

  // NEW METHOD: Get user profiles with site access
  private async getUserProfilesWithAccess(siteId: string): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      // This method tries to find users who have accessed the site recently
      const client = await this.getGraphClient();
      
      // Try to get site analytics to find active users
      try {
        const analyticsResponse = await client
          .api(`/sites/${siteId}/analytics/allTime`)
          .get();
        
        // This endpoint might give us insights into site usage
        console.log('Site analytics available:', analyticsResponse);
        
      } catch (error) {
        console.debug('Site analytics not available:', error);
      }
      
      // Try to get recent activities
      try {
        const activitiesResponse = await client
          .api(`/sites/${siteId}/drive/recent`)
          .get();
        
        // Extract user information from recent activities
        for (const item of activitiesResponse.value || []) {
          const activity = item as Record<string, unknown>;
          if (activity.createdBy || activity.lastModifiedBy) {
            const userInfo = activity.createdBy || activity.lastModifiedBy;
            const user = (userInfo as Record<string, unknown>)?.user as Record<string, unknown>;
            
            if (user && user.id) {
              allUsers.push({
                id: user.id as string,
                displayName: user.displayName as string || 'Recent User',
                mail: user.email as string || '',
                userPrincipalName: user.userPrincipalName as string || user.email as string || '',
                accessLevel: 'member' as const,
                source: 'site' as const
              });
            }
          }
        }
        
      } catch (error) {
        console.debug('Recent activities not available:', error);
      }
      
    } catch (error) {
      console.debug('User profiles with access retrieval failed:', error);
    }
    
    return allUsers;
  }

  // NEW METHOD: Get ALL site users using SharePoint's native user enumeration (like _layouts/15/user.aspx)
  private async getAllSiteUsersFromSharePoint(): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      const spHttpClient = this.context.spHttpClient;
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Method 1: Get ALL site users with detailed information
      const detailedUsersResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/siteusers?$select=Id,Title,Email,UserPrincipalName,LoginName,IsSiteAdmin,PrincipalType&$filter=PrincipalType eq 1`,
        SPHttpClient.configurations.v1
      );
      
      if (detailedUsersResponse.ok) {
        const detailedUsersData = await detailedUsersResponse.json();
        const users = detailedUsersData.d?.results || [];
        
        for (const user of users) {
          // Skip system accounts but include all real users
          if (user.LoginName && 
              !user.LoginName.includes('c:0(.s|true)') && 
              !user.LoginName.includes('SHAREPOINT\\\\system') &&
              !user.LoginName.includes('app@sharepoint') &&
              user.PrincipalType === 1 && // User principal type
              user.Email) {
            
            // Determine access level based on IsSiteAdmin and other factors
            let accessLevel: 'owner' | 'admin' | 'member' | 'visitor' = 'member';
            if (user.IsSiteAdmin) {
              accessLevel = 'owner';
            }
            
            allUsers.push({
              id: user.LoginName || user.Id?.toString() || 'unknown',
              displayName: user.Title || user.Email || 'Unknown User',
              mail: user.Email || '',
              userPrincipalName: user.UserPrincipalName || user.Email || user.LoginName,
              accessLevel,
              source: 'site' as const
            });
          }
        }
        
        console.log(`SharePoint native enumeration found ${allUsers.length} users`);
      }
      
      // Method 2: Get users from all web role assignments (this is the most comprehensive)
      const roleAssignmentsResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Title,Member/Email,Member/UserPrincipalName,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/BasePermissions`,
        SPHttpClient.configurations.v1
      );
      
      if (roleAssignmentsResponse.ok) {
        const roleData = await roleAssignmentsResponse.json();
        const assignments = roleData.d?.results || [];
        
        for (const assignment of assignments) {
          const member = assignment.Member;
          const roleBindings = assignment.RoleDefinitionBindings?.results || [];
          
          if (member && member.PrincipalType === 1 && member.Email) { // User principal type
            // Determine access level based on role bindings
            let accessLevel: 'owner' | 'admin' | 'member' | 'visitor' = 'visitor';
            
            for (const role of roleBindings) {
              const roleName = (role.Name || '').toLowerCase();
              if (roleName.includes('full control') || roleName.includes('owner')) {
                accessLevel = 'owner';
                break;
              } else if (roleName.includes('design') || roleName.includes('manage')) {
                accessLevel = 'admin';
              } else if (roleName.includes('contribute') || roleName.includes('edit')) {
                accessLevel = 'member';
              } else if (roleName.includes('read') || roleName.includes('view')) {
                accessLevel = 'visitor';
              }
            }
            
            allUsers.push({
              id: member.LoginName || 'unknown',
              displayName: member.Title || member.Email || 'Unknown User',
              mail: member.Email || '',
              userPrincipalName: member.UserPrincipalName || member.Email || member.LoginName,
              accessLevel,
              source: 'site' as const
            });
          }
        }
        
        console.log(`Role assignments found ${assignments.length} total assignments`);
      }
      
    } catch (error) {
      console.warn('SharePoint native user enumeration failed:', error);
    }
    
    return allUsers;
  }

  // NEW METHOD: Get users from role assignments (replicates _layouts/15/user.aspx functionality)
  private async getUsersFromRoleAssignments(): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      const spHttpClient = this.context.spHttpClient;
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Get all role assignments for the web, including groups
      const roleAssignmentsResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/roleassignments?$expand=Member,Member/Users,RoleDefinitionBindings&$select=Member/Id,Member/Title,Member/LoginName,Member/Email,Member/UserPrincipalName,Member/PrincipalType,Member/Users/Title,Member/Users/Email,Member/Users/UserPrincipalName,Member/Users/LoginName,RoleDefinitionBindings/Name`,
        SPHttpClient.configurations.v1
      );
      
      if (roleAssignmentsResponse.ok) {
        const roleData = await roleAssignmentsResponse.json();
        const assignments = roleData.d?.results || [];
        
        for (const assignment of assignments) {
          const member = assignment.Member;
          const roleBindings = assignment.RoleDefinitionBindings?.results || [];
          
          // Determine access level from role bindings
          let accessLevel: 'owner' | 'admin' | 'member' | 'visitor' = 'visitor';
          for (const role of roleBindings) {
            const roleName = (role.Name || '').toLowerCase();
            if (roleName.includes('full control') || roleName.includes('owner')) {
              accessLevel = 'owner';
              break;
            } else if (roleName.includes('design') || roleName.includes('manage') || roleName.includes('contribute')) {
              accessLevel = accessLevel === 'visitor' ? 'member' : accessLevel;
            } else if (roleName.includes('read') || roleName.includes('view')) {
              accessLevel = accessLevel === 'visitor' ? 'visitor' : accessLevel;
            }
          }
          
          if (member.PrincipalType === 1 && member.Email) {
            // Direct user assignment
            allUsers.push({
              id: member.LoginName || member.Id?.toString() || 'unknown',
              displayName: member.Title || member.Email || 'Unknown User',
              mail: member.Email || '',
              userPrincipalName: member.UserPrincipalName || member.Email || member.LoginName,
              accessLevel,
              source: 'site' as const
            });
          } else if (member.PrincipalType === 8 && member.Users) {
            // SharePoint group - get all users from the group
            const groupUsers = member.Users.results || [];
            for (const user of groupUsers) {
              if (user.Email) {
                allUsers.push({
                  id: user.LoginName || 'unknown',
                  displayName: user.Title || user.Email || 'Unknown User',
                  mail: user.Email || '',
                  userPrincipalName: user.UserPrincipalName || user.Email || user.LoginName,
                  accessLevel,
                  source: 'site' as const
                });
              }
            }
          }
        }
        
        console.log(`Role assignments method found ${assignments.length} role assignments`);
      }
      
    } catch (error) {
      console.warn('Role assignments user retrieval failed:', error);
    }
    
    return allUsers;
  }

  // SIMPLE BACKUP METHOD: Get SharePoint users with minimal configuration
  private async getSimpleSharePointUsers(): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    try {
      const spHttpClient = this.context.spHttpClient;
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      
      console.log('Trying simple SharePoint REST API calls...');
      
      // Try the most basic site users call without complex selects and with basic headers
      const basicUsersResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/siteusers`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json'
          }
        }
      );
      
      if (basicUsersResponse.ok) {
        const basicUsersData = await basicUsersResponse.json();
        const users = basicUsersData.value || basicUsersData.d?.results || [];
        
        console.log(`Found ${users.length} basic site users`);
        
        for (const user of users) {
          // Include all users that have an email and are not system accounts
          const upn = user.UserPrincipalName || user.Email || user.LoginName;
          const email = user.Email;
          
          if (email && 
              user.LoginName && 
              upn &&
              !user.LoginName.includes('SHAREPOINT\\\\system') &&
              !user.LoginName.includes('app@sharepoint') &&
              !upn.includes('app@sharepoint') &&
              !email.includes('noreply') &&
              user.Title !== 'System Account' &&
              user.Title !== 'SharePoint App' &&
              user.PrincipalType === 1) { // Ensure it's a user, not a group
            
            allUsers.push({
              id: user.LoginName || user.Id?.toString() || 'unknown',
              displayName: user.Title || email || 'User',
              mail: email,
              userPrincipalName: upn,
              accessLevel: user.IsSiteAdmin ? 'owner' : 'member' as const,
              source: 'site' as const
            });
          }
        }
      } else {
        console.warn('Basic site users call failed:', basicUsersResponse.status, basicUsersResponse.statusText);
      }
      
      // Try basic site groups call
      const basicGroupsResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/sitegroups`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json'
          }
        }
      );
      
      if (basicGroupsResponse.ok) {
        const basicGroupsData = await basicGroupsResponse.json();
        const groups = basicGroupsData.value || basicGroupsData.d?.results || [];
        
        console.log(`Found ${groups.length} basic site groups`);
        
        // For each group, try to get its users
        for (const group of groups) {
          try {
            const groupUsersResponse = await spHttpClient.get(
              `${siteUrl}/_api/web/sitegroups(${group.Id})/users`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  'Accept': 'application/json'
                }
              }
            );
            
            if (groupUsersResponse.ok) {
              const groupUsersData = await groupUsersResponse.json();
              const groupUsers = groupUsersData.value || groupUsersData.d?.results || [];
              
              // Determine access level based on group title
              let accessLevel: 'owner' | 'admin' | 'member' | 'visitor' = 'visitor';
              const groupTitle = (group.Title || '').toLowerCase();
              if (groupTitle.includes('owner')) {
                accessLevel = 'owner';
              } else if (groupTitle.includes('member') || groupTitle.includes('contribute')) {
                accessLevel = 'member';
              } else if (groupTitle.includes('visitor') || groupTitle.includes('read')) {
                accessLevel = 'visitor';
              }
              
              for (const user of groupUsers) {
                const upn = user.UserPrincipalName || user.Email || user.LoginName;
                const email = user.Email;
                
                if (email && 
                    user.LoginName && 
                    upn &&
                    !user.LoginName.includes('SHAREPOINT\\\\system') &&
                    !upn.includes('app@sharepoint') &&
                    !email.includes('noreply') &&
                    user.Title !== 'System Account') {
                    
                  allUsers.push({
                    id: user.LoginName || 'unknown',
                    displayName: user.Title || email || 'User',
                    mail: email,
                    userPrincipalName: upn,
                    accessLevel,
                    source: 'site' as const
                  });
                }
              }
            }
          } catch (error) {
            console.debug(`Could not get users for group ${group.Title}:`, error);
          }
        }
      } else {
        console.warn('Basic site groups call failed:', basicGroupsResponse.status, basicGroupsResponse.statusText);
      }
      
    } catch (error) {
      console.warn('Simple SharePoint users backup method failed:', error);
    }
    
    return allUsers;
  }

  private async getFallbackSiteMembers(siteId: string): Promise<IUser[]> {
    const allUsers: IUser[] = [];
    
    // Try multiple fallback approaches
    const fallbackMethods = [
      // Method 1: Try getting site information with creator
      async () => {
        try {
          const client = await this.getGraphClient();
          const response = await client
            .api(`/sites/${siteId}`)
            .select('createdBy')
            .get();
          
          const r = response as Record<string, unknown>;
          const createdBy = r.createdBy as Record<string, unknown>;
          
          if (createdBy && createdBy.user) {
            const user = createdBy.user as Record<string, unknown>;
            return [{
              id: user.id as string || 'creator',
              displayName: user.displayName as string || 'Site Creator',
              mail: user.email as string,
              userPrincipalName: user.email as string,
              accessLevel: 'owner' as const,
              source: 'site' as const
            }];
          }
          return [];
        } catch (error) {
          console.log('Site creator method failed:', error);
          return [];
        }
      },
      
      // Method 2: Try getting basic site information
      async () => {
        try {
          const client = await this.getGraphClient();
          await client
            .api(`/sites/${siteId}`)
            .get();
          
          // At minimum, we can show that the site exists and has the current user
          console.log('Site basic info retrieved, but no specific member information available');
          return [];
        } catch (error) {
          console.log('Site basic info method failed:', error);
          return [];
        }
      }
    ];
    
    // Try each fallback method
    for (const method of fallbackMethods) {
      try {
        const users = await method();
        if (users.length > 0) {
          allUsers.push(...users);
          console.log(`Fallback method found ${users.length} users`);
        }
      } catch (error) {
        console.log('Fallback method failed:', error);
        continue;
      }
    }
    
    if (allUsers.length === 0) {
      console.error('All fallback site member methods failed');
      throw new Error('No fallback methods could retrieve site members');
    }
    
    return allUsers;
  }
}