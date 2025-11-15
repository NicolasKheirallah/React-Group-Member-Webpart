import * as React from 'react';
import { useEffect, useCallback, useMemo, useState } from 'react';
import {
  Persona,
  PersonaSize,
  Spinner,
  SpinnerSize,
  DefaultButton,
  PersonaInitialsColor,
  Text,
  Stack,
  StackItem,
  SearchBox,
  IconButton,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  FocusZone,
  List,
  CommandBar,
  ICommandBarItemProps,
  Pivot,
  PivotItem,
  Separator,
  Shimmer,
  ActionButton
} from '@fluentui/react';
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";
import { IUser, IUsersByRole, UserPersonaProps } from '../types/interfaces';
import { IGroupMembersProps } from './IGroupMembersProps';
import styles from './GroupMembers.module.scss';
import UnifiedProfileImage from './UnifiedProfileImage';
import ErrorBoundary from './ErrorBoundary';

// Import the new architecture
import { 
  useUsers, 
  useSearch, 
  useLoadingState, 
  usePaginatedUsers,
  usePresence 
} from '../hooks/useStateManager';
import { 
  useUnifiedGraphService, 
  useLoggingService
} from '../services/ServiceContainer';

type AccessRole = 'owner' | 'admin' | 'member' | 'visitor';

const getFallbackInitials = (displayName: string): string => {
  const names = displayName.trim().split(' ');
  if (names.length === 1) {
    return names[0].charAt(0).toUpperCase();
  }
  const firstLetter = names[0].charAt(0).toUpperCase();
  const lastLetter = names[names.length - 1].charAt(0).toUpperCase();
  return `${firstLetter}${lastLetter}`;
};

interface UserPersonaWithServiceProps extends UserPersonaProps {
  presenceEnabled: boolean;
  showRoleLabels: boolean;
  roleLabels: Record<AccessRole, string>;
}

const UserPersona: React.FC<UserPersonaWithServiceProps> = React.memo(({ user, presenceEnabled, showRoleLabels, roleLabels }) => {
  const graphService = useUnifiedGraphService();
  
  // Guard against invalid user data
  if (!user || !user.displayName || !user.id) {
    return null;
  }
  
  const fallbackInitials = getFallbackInitials(user.displayName);

  const roleText = showRoleLabels && user.accessLevel ? roleLabels[user.accessLevel as AccessRole] || user.accessLevel : user.jobTitle || 'Member';

  return (
    <Persona
      text={user.displayName}
      secondaryText={roleText}
      tertiaryText={user.department}
      optionalText={user.officeLocation}
      size={PersonaSize.size40}
      initialsColor={PersonaInitialsColor.blue}
      imageInitials={fallbackInitials}
      onRenderPersonaCoin={() => (
        <UnifiedProfileImage
          userId={user.id}
          userPrincipalName={user.userPrincipalName}
          graphService={graphService}
          fallbackInitials={fallbackInitials}
          alt={user.displayName}
          className={styles.profileImage}
          showPresence={presenceEnabled}
        />
      )}
    />
  );
});

const roleLabelMap = (props: IGroupMembersProps): Record<AccessRole, string> => ({
  owner: props.ownerLabel || 'Owners',
  admin: props.adminLabel || 'Administrators',
  member: props.memberLabel || 'Members',
  visitor: props.visitorLabel || 'Visitors'
});

const claimPrincipalPatterns = ['c:0', 'spo-grid-all-users', 'everyone except external users', 'everyone'];

const normalizePatternList = (patterns?: string): string[] => {
  if (!patterns) {
    return [];
  }
  return patterns
    .split(/\r?\n/)
    .map(entry => entry.trim().toLowerCase())
    .filter(entry => entry.length > 0);
};

const defaultExcludedPrincipals = [
  'sharepoint\\system',
  'sharepoint\\app',
  'nt service\\',
  'app@sharepoint'
];

const EnhancedGroupMembers: React.FC<IGroupMembersProps> = (props): JSX.Element => {
  const { loading, error, actions: userActions, selectors } = useUsers();
  const { searchTerm, searchResults, setSearchTerm, clearSearch } = useSearch();
  const { retry } = useLoadingState();
  const { presenceEnabled, togglePresence } = usePresence();

  const graphService = useUnifiedGraphService();
  const logger = useLoggingService();
  const [activeRoleKey, setActiveRoleKey] = useState<string>('all');
  const availableRoles = useMemo(() => (props.roles.length ? props.roles : ['member']), [props.roles]);
  const labels = useMemo(() => roleLabelMap(props), [props.ownerLabel, props.adminLabel, props.memberLabel, props.visitorLabel]);
  const showSummaryGrid = props.showSummaryGrid ?? false;
  const showRolePivot = props.showRolePivot ?? false;
  const showPageHeader = props.showPageHeader ?? false;
  const showCommandBar = props.showCommandBar !== false;
  const showSectionBorders = props.showSectionBorders !== false;
  const customHeaderTitle = props.pageHeaderTitle || 'People directory';
  const customHeaderSubtitle = props.pageHeaderSubtitle || (searchResults.isFiltered
    ? `Showing ${searchResults.resultCount} of ${searchResults.totalCount}`
    : `${searchResults.totalCount} people in this site`);

  const cacheKey = useMemo(() => `gmw-exclusions-${props.context.instanceId}`, [props.context.instanceId]);

  const exclusionPatterns = useMemo(() => {
    const custom = normalizePatternList(props.excludedPrincipals);
    const signature = `${props.hideClaimsPrincipals !== false ? 'claims-on' : 'claims-off'}|${custom.join('|')}`;
    if (typeof sessionStorage !== 'undefined') {
      try {
        const stored = sessionStorage.getItem(cacheKey);
        if (stored) {
          const parsed = JSON.parse(stored) as { signature: string; patterns: string[] };
          if (parsed.signature === signature) {
            return parsed.patterns;
          }
        }
      } catch {
        // ignore
      }
    }

    const basePatterns = [...defaultExcludedPrincipals];
    if (props.hideClaimsPrincipals !== false) {
      basePatterns.push(...claimPrincipalPatterns);
    }

    const combined = Array.from(new Set([...basePatterns, ...custom].map(pattern => pattern.toLowerCase())));
    if (typeof sessionStorage !== 'undefined') {
      try {
        sessionStorage.setItem(cacheKey, JSON.stringify({ signature, patterns: combined }));
      } catch {
        // ignore storage errors
      }
    }
    return combined;
  }, [props.excludedPrincipals, cacheKey, props.hideClaimsPrincipals]);

  const isClaimsPrincipal = useCallback((user: IUser): boolean => {
    const candidates = [user.userPrincipalName, user.mail, user.id, user.displayName]
      .map(value => value?.toLowerCase())
      .filter(Boolean) as string[];
    return candidates.some(candidate => claimPrincipalPatterns.some(pattern => candidate.includes(pattern)));
  }, []);

  const shouldExcludeUser = useCallback((user: IUser): boolean => {
    if (props.hideClaimsPrincipals !== false && isClaimsPrincipal(user)) {
      return true;
    }
    if (!exclusionPatterns.length) {
      return false;
    }
    const candidates = [user.userPrincipalName, user.mail, user.id, user.displayName]
      .map(value => value?.toLowerCase())
      .filter(Boolean) as string[];
    return candidates.some(candidate => exclusionPatterns.some(pattern => candidate.includes(pattern)));
  }, [exclusionPatterns, isClaimsPrincipal, props.hideClaimsPrincipals]);

  const isGroupPrincipal = useCallback((user: IUser): boolean => {
    if (user.isGroup || user.principalType?.toLowerCase().includes('group')) {
      return true;
    }
    return isClaimsPrincipal(user);
  }, [isClaimsPrincipal]);

  useEffect(() => {
    setActiveRoleKey('all');
  }, [availableRoles]);

  const rolePriority: Record<AccessRole, number> = {
    owner: 4,
    admin: 3,
    member: 2,
    visitor: 1
  };

  const normalize = (value?: string): string | undefined => value?.trim().toLowerCase() || undefined;

  const getUserKey = (user: IUser): string | undefined => {
    return (
      normalize(user.userPrincipalName) ||
      normalize(user.mail) ||
      normalize(user.id)
    );
  };

  const fetchGroupUsers = useCallback(async (): Promise<void> => {
    userActions.loadUsersStart();

    const timerId = logger.startTimer('fetchAllSiteMembers');

    try {
      logger.info('GroupMembers', 'Starting to fetch site members', { roles: props.roles });

      const allMembers = (await graphService.getAllSiteMembers()).filter(user => !shouldExcludeUser(user));

      const dedupedUsers = new Map<string, IUser>();

      for (const user of allMembers) {
        const key = getUserKey(user);
        if (!key) {
          continue;
        }

        const accessLevel = (user.accessLevel || 'visitor') as AccessRole;
        const existing = dedupedUsers.get(key);

        if (!existing) {
          dedupedUsers.set(key, user);
        } else {
          const existingAccess = (existing.accessLevel || 'visitor') as AccessRole;
          if (rolePriority[accessLevel] > rolePriority[existingAccess]) {
            dedupedUsers.set(key, { ...user, accessLevel });
          }
        }
      }

      const newUsersByRole: IUsersByRole = {
        owner: [],
        admin: [],
        member: [],
        visitor: []
      };

      dedupedUsers.forEach(user => {
        const accessLevel = (user.accessLevel || 'visitor') as AccessRole;
        if (props.roles.includes(accessLevel as string)) {
          newUsersByRole[accessLevel].push(user);
        }
      });

      userActions.loadUsersSuccess(newUsersByRole);
      
      logger.endTimer(timerId, 'GroupMembers', 'Successfully fetched site members', {
        totalUsers: allMembers.length,
        roles: Object.keys(newUsersByRole).map(role => ({
          role,
          count: newUsersByRole[role as keyof IUsersByRole].length
        }))
      });
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Error retrieving site members.";
      userActions.loadUsersError(errorMessage);
      
      logger.endTimer(timerId, 'GroupMembers', 'Failed to fetch site members', { error: errorMessage });
      logger.error('GroupMembers', 'Error in fetchGroupUsers', error as Error);
    }
  }, [graphService, props.roles, userActions, logger]);

  const handleSearchChange = useCallback((newValue?: string): void => {
    setSearchTerm(newValue || "");
  }, [setSearchTerm]);

  useEffect(() => {
    fetchGroupUsers().catch((error: Error) => {
      logger.error('GroupMembers', 'Failed to fetch group users on mount', error);
    });
  }, [fetchGroupUsers, logger]);

  const UserSection: React.FC<{ role: keyof IUsersByRole }> = ({ role }): JSX.Element | null => {
    const { users, totalPages, currentPage, hasMore, actions } = usePaginatedUsers(role);

    useEffect(() => {
      if (!users.length) {
        return;
      }
      const userIds = users
        .map(user => user.id)
        .filter((id): id is string => typeof id === 'string' && id.length > 0);

      graphService.getBatchUserPresence(userIds).catch(error => {
        logger.warn('GroupMembers', 'Failed to batch presence', error);
      });

      graphService.prefetchUserPhotos(users).catch(error => {
        logger.warn('GroupMembers', 'Failed to prefetch photos', error);
      });
    }, [users, graphService, logger]);
    
    if (users.length === 0 && !searchResults.isFiltered) {
      return null;
    }
    
    const sectionClassName = [styles.userSection, showSectionBorders ? '' : styles.userSectionBorderless]
      .filter(Boolean)
      .join(' ');

    return (
      <ErrorBoundary context={`UserSection-${role}`} level="component">
        <div className={sectionClassName}>
          <Stack horizontal verticalAlign="center" className={styles.sectionHeader}>
            <Text variant="large" as="h3" className={styles.sectionTitle}>
              {labels[role as AccessRole]} ({users.length}{searchResults.isFiltered ? ` of ${searchResults.resultCount}` : ''})
            </Text>
            <StackItem grow>
              <div className={styles.sectionDivider} />
            </StackItem>
          </Stack>
          
          <div className={styles.userList}>
            <FocusZone>
              <List
                items={users}
                onRenderCell={(user: IUser | undefined): JSX.Element | null => {
                  if (!user || !user.displayName || !user.id) return null;
                  return (
                    <ErrorBoundary context={`UserListItem-${user.id}`} level="component">
                      <div className={styles.listItem}>
                        <div className={styles.personaContainer}>
                          <div className={styles.personaWrapper}>
                            <LivePersona
                              upn={user.userPrincipalName}
                              serviceScope={props.context.serviceScope}
                              template={
                                <UserPersona
                                  user={user}
                                  context={props.context}
                                  presenceEnabled={presenceEnabled}
                                  showRoleLabels={props.showRoleLabels ?? false}
                                  roleLabels={labels}
                                />
                              }
                            />
                            {isGroupPrincipal(user) && (
                              <span className={styles.principalBadge}>Group</span>
                            )}
                          </div>
                        </div>
                        <div className={styles.listActions}>
                          <IconButton
                            iconProps={{ iconName: 'Chat' }}
                            title={`Start a chat with ${user.displayName}`}
                            ariaLabel={`Start a chat with ${user.displayName}`}
                            onClick={(e) => {
                              e.stopPropagation();
                              logger.info('GroupMembers', 'Starting Teams chat', { userId: user.id, userPrincipalName: user.userPrincipalName });
                              window.open(`https://teams.microsoft.com/l/chat/0/0?users=${user.userPrincipalName}`, '_blank');
                            }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Mail' }}
                            title={`Send email to ${user.displayName}`}
                            ariaLabel={`Send email to ${user.displayName}`}
                            onClick={(e) => {
                              e.stopPropagation();
                              logger.info('GroupMembers', 'Opening email client', { userId: user.id, mail: user.mail });
                              window.location.href = `mailto:${user.mail}`;
                            }}
                          />
                        </div>
                      </div>
                    </ErrorBoundary>
                  );
                }}
              />
            </FocusZone>
          </div>
          
          {totalPages > 1 && (
            <div className={styles.paginationContainer}>
              <div className={styles.paginationControls}>
                <ActionButton
                  className={styles.paginationButton}
                  iconProps={{ iconName: 'ChevronLeft' }}
                  disabled={currentPage === 1}
                  onClick={() => actions.prevPage()}
                  text="Previous"
                />
                <Text variant="medium" className={styles.paginationText}>
                  Page {currentPage} of {totalPages}
                </Text>
                <ActionButton
                  className={styles.paginationButton}
                  iconProps={{ iconName: 'ChevronRight' }}
                  disabled={!hasMore}
                  onClick={() => actions.nextPage()}
                  text="Next"
                />
              </div>
            </div>
          )}
        </div>
      </ErrorBoundary>
    );
  };

  const summaryData = useMemo(() => availableRoles.map(role => {
    const filteredUsers = selectors.getFilteredUsers(role as keyof IUsersByRole);
    return {
      role,
      label: labels[role as AccessRole],
      count: filteredUsers.length
    };
  }), [availableRoles, labels, selectors]);

  const pivotItems = useMemo(() => ['all', ...availableRoles], [availableRoles]);
  const displayedRoles = useMemo(() => {
    if (!showRolePivot || activeRoleKey === 'all') {
      return availableRoles;
    }
    return availableRoles.filter(role => role === activeRoleKey);
  }, [activeRoleKey, availableRoles, showRolePivot]);

  const commandItems: ICommandBarItemProps[] = useMemo(() => [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        fetchGroupUsers().catch(console.error);
      }
    },
    {
      key: 'presence',
      text: presenceEnabled ? 'Hide presence' : 'Show presence',
      iconProps: { iconName: 'PresenceChickletVideo' },
      onClick: () => togglePresence(!presenceEnabled)
    }
  ], [fetchGroupUsers, presenceEnabled, togglePresence]);

  const farItems: ICommandBarItemProps[] = useMemo(() => (
    searchResults.isFiltered ? [{
      key: 'clearSearch',
      text: 'Clear search',
      iconProps: { iconName: 'Clear' },
      onClick: clearSearch
    }] : []
  ), [searchResults.isFiltered, clearSearch]);

  return (
    <ErrorBoundary context="GroupMembers" level="page">
      <div className={styles.groupMembers}>
        <div className={styles.pageHeader}>
          {showPageHeader && (
            <div className={styles.headerText}>
              <Text variant="xLarge" className={styles.title}>{customHeaderTitle}</Text>
              <Text variant="smallPlus" className={styles.subtitle}>{customHeaderSubtitle}</Text>
            </div>
          )}
          {showCommandBar && (
            <CommandBar items={commandItems} farItems={farItems} className={styles.commandBar} ariaLabel="People actions" />
          )}
        </div>

        {showSummaryGrid && (
          <div className={styles.summaryGrid}>
            {summaryData.map(summary => (
              <div key={summary.role} className={styles.summaryCard}>
                <Text variant="smallPlus" className={styles.summaryLabel}>{summary.label}</Text>
                <Text variant="xxLarge" className={styles.summaryValue}>{summary.count}</Text>
              </div>
            ))}
          </div>
        )}

        {props.showSearchBox && (
          <div className={styles.searchContainer}>
            <SearchBox
              placeholder="Search by name, role, department, or location..."
              onChange={(_, newValue) => handleSearchChange(newValue)}
              iconProps={{ iconName: 'Search' }}
              className={styles.searchBox}
              value={searchTerm}
            />
          </div>
        )}

        {showRolePivot && (
          <div className={styles.pivotContainer}>
            <Pivot selectedKey={activeRoleKey} onLinkClick={(item) => setActiveRoleKey(item?.props.itemKey || 'all')}>
              <PivotItem headerText="All" itemKey="all" />
              {pivotItems.filter(item => item !== 'all').map(role => (
                <PivotItem key={role} headerText={labels[role as AccessRole]} itemKey={role} />
              ))}
            </Pivot>
          </div>
        )}

        {(showRolePivot || showSummaryGrid || showPageHeader) && <Separator className={styles.separator} />}

        {loading && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Loading group users..." />
            <Shimmer width="90%" />
            <Shimmer width="75%" />
            <ProgressIndicator label="Retrieving user information" description="Please wait..." />
          </div>
        )}

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={true}
            dismissButtonAriaLabel="Close"
            onDismiss={() => userActions.loadUsersError('')}
            className={styles.errorMessage}
            actions={
              <div>
                <DefaultButton
                  onClick={() => {
                    retry();
                    fetchGroupUsers().catch(console.error);
                  }}
                  text="Retry"
                  iconProps={{ iconName: 'Refresh' }}
                />
              </div>
            }
          >
            <strong>Failed to load group members</strong>
            <br />
            {error}
            <br />
            <small>
              Please check your permissions and network connection. If the problem persists, contact your administrator.
            </small>
          </MessageBar>
        )}

        {!loading && !error && (
          <div className={styles.contentContainer}>
            {displayedRoles.map(role => (
              <UserSection
                key={role}
                role={role as keyof IUsersByRole}
              />
            ))}
            {searchResults.isFiltered && !searchResults.hasResults && (
              <MessageBar messageBarType={MessageBarType.info}>
                <Text>No users found matching &quot;{searchTerm}&quot;. Try adjusting your search terms.</Text>
              </MessageBar>
            )}
          </div>
        )}
      </div>
    </ErrorBoundary>
  );
};

export default EnhancedGroupMembers;
