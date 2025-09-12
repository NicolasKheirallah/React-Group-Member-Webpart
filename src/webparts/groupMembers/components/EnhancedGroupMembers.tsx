import * as React from 'react';
import { useEffect, useCallback } from 'react';
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
  PrimaryButton,
  SearchBox,
  IconButton,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  FocusZone,
  List
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
}

const UserPersona: React.FC<UserPersonaWithServiceProps> = React.memo(({ user, presenceEnabled }) => {
  const graphService = useUnifiedGraphService();
  
  // Guard against invalid user data
  if (!user || !user.displayName || !user.id) {
    return null;
  }
  
  const fallbackInitials = getFallbackInitials(user.displayName);

  return (
    <Persona
      text={user.displayName}
      secondaryText={user.jobTitle || 'Member'}
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

const EnhancedGroupMembers: React.FC<IGroupMembersProps> = (props): JSX.Element => {
  const { loading, error, actions: userActions } = useUsers();
  const { searchTerm, searchResults, setSearchTerm, clearSearch } = useSearch();
  const { retry } = useLoadingState();
  const { presenceEnabled } = usePresence();
  
  const graphService = useUnifiedGraphService();
  const logger = useLoggingService();

  const fetchGroupUsers = useCallback(async (): Promise<void> => {
    userActions.loadUsersStart();
    
    const timerId = logger.startTimer('fetchAllSiteMembers');
    
    try {
      logger.info('GroupMembers', 'Starting to fetch site members', { roles: props.roles });
      
      const allMembers = await graphService.getAllSiteMembers();
      
      const newUsersByRole: IUsersByRole = {
        owner: [],
        admin: [],
        member: [],
        visitor: []
      };

      for (const user of allMembers) {
        const accessLevel = user.accessLevel || 'visitor';
        if (props.roles.includes(accessLevel)) {
          newUsersByRole[accessLevel].push(user);
        }
      }

      for (const role of props.roles) {
        if (!newUsersByRole[role as keyof IUsersByRole]) {
          newUsersByRole[role as keyof IUsersByRole] = [];
        }
      }

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
    
    if (users.length === 0 && !searchResults.isFiltered) {
      return null;
    }
    
    const roleLabels: Record<string, string> = {
      owner: props.ownerLabel || 'Owners',
      admin: props.adminLabel || 'Administrators', 
      member: props.memberLabel || 'Members',
      visitor: props.visitorLabel || 'Visitors'
    };

    const handleLoadMore = (): void => {
      actions.nextPage();
    };

    return (
      <ErrorBoundary context={`UserSection-${role}`} level="component">
        <div className={styles.userSection}>
          <Stack horizontal verticalAlign="center" className={styles.sectionHeader}>
            <Text variant="large" as="h3" className={styles.sectionTitle}>
              {roleLabels[role]} ({users.length}{searchResults.isFiltered ? ` of ${searchResults.resultCount}` : ''})
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
                          <LivePersona
                            upn={user.userPrincipalName}
                            serviceScope={props.context.serviceScope}
                            template={
                              <UserPersona
                                user={user}
                                context={props.context}
                                presenceEnabled={presenceEnabled}
                              />
                            }
                          />
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
              {hasMore ? (
                <PrimaryButton
                  text="Load More"
                  onClick={handleLoadMore}
                  className={styles.loadMoreButton}
                  iconProps={{ iconName: 'ChevronDown' }}
                />
              ) : (
                <div className={styles.paginationControls}>
                  <DefaultButton
                    text="Previous"
                    onClick={() => actions.prevPage()}
                    disabled={currentPage === 1}
                    iconProps={{ iconName: 'ChevronLeft' }}
                  />
                  <Text variant="medium" className={styles.paginationText}>
                    Page {currentPage} of {totalPages}
                  </Text>
                  <DefaultButton
                    text="Next"
                    onClick={() => actions.nextPage()}
                    disabled={!hasMore}
                    iconProps={{ iconName: 'ChevronRight' }}
                  />
                </div>
              )}
            </div>
          )}
        </div>
      </ErrorBoundary>
    );
  };

  return (
    <ErrorBoundary context="GroupMembers" level="page">
      <div className={styles.groupMembers}>
        {props.showSearchBox && (
          <div className={styles.searchContainer}>
            <SearchBox
              placeholder="Search by name, role, department, or location..."
              onChange={(_, newValue) => handleSearchChange(newValue)}
              iconProps={{ iconName: 'Search' }}
              className={styles.searchBox}
              underlined
              value={searchTerm}
            />
            {searchResults.isFiltered && (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
                <Text variant="small">
                  Showing {searchResults.resultCount} of {searchResults.totalCount} users
                </Text>
                <DefaultButton
                  text="Clear"
                  iconProps={{ iconName: 'Clear' }}
                  onClick={clearSearch}
                  styles={{ root: { minWidth: 'auto' } }}
                />
              </Stack>
            )}
          </div>
        )}
        
        {loading && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Loading group users..." />
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
            {props.roles.map(role => (
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