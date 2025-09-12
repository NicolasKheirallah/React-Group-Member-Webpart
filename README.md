# Group Members Web Part

A modern SharePoint Framework web part for displaying and managing site members from both Microsoft 365 groups and SharePoint Communication sites with advanced features and responsive design.

## Overview

This web part provides a comprehensive solution for viewing and interacting with site members across different SharePoint environments. Built with modern React patterns and Microsoft Graph integration, it automatically detects whether you're on a Microsoft 365 group-connected site or a Communication site and shows all users with appropriate access levels.

### Key Features

- **Unified Site Member Discovery**: Works with both M365 Groups and Communication sites
- **Smart Access Level Detection**: Automatically identifies Owners, Administrators, Members, and Visitors
- **Microsoft Graph Integration**: Real-time data from Microsoft 365 and SharePoint APIs
- **Smart Profile Images**: Automated fallback with initials, intelligent caching
- **Advanced Search**: Debounced search across multiple user properties
- **Role-based Filtering**: Owners, Administrators, Members, and Visitors
- **Interactive Actions**: Direct Teams chat and email integration
- **Customizable Title**: SharePoint native title editing with inline editing support
- **Responsive Design**: Optimized for desktop and mobile experiences
- **Accessibility**: Full screen reader and keyboard navigation support

## Features

- **Intelligent Site Detection**: Automatically detects M365 Groups vs Communication sites
- **Unified Member Access**: 
  - M365 Group members (owners, admins, members)
  - SharePoint site permissions (contribute, read, full control)
  - Unified access level mapping (owner/admin/member/visitor)
- **Custom Profile Image Handling**: 
  - Graceful fallback for profile photos
  - Lazy loading of images
  - Initials-based avatar generation
- **Flexible Configuration**: 
  - Configurable roles (Owner, Admin, Member, Visitor)
  - Customizable labels
  - Pagination settings
- **Enhanced User Experience**:
  - Search functionality
  - Quick action buttons
  - Customizable web part title (inline editing)
  - Responsive design

## Web Part Properties

| Property | Type | Description | Default | Required |
|----------|------|-------------|---------|----------|
| `title` | String | Custom title for the web part | 'Site Members' | No |
| `roles` | Array of Strings | Roles to display (owner, admin, member, visitor) | None | Yes |
| `itemsPerPage` | Number | Number of users per page | 10 | No |
| `sortField` | String | Sort users by 'name' or 'jobTitle' | 'name' | No |
| `showSearchBox` | Boolean | Enable/disable search functionality | true | No |
| `showPresenceIndicator` | Boolean | Show Microsoft Teams presence status | true | No |
| `ownerLabel` | String | Custom label for owners | 'Owners' | No |
| `adminLabel` | String | Custom label for administrators | 'Administrators' | No |
| `memberLabel` | String | Custom label for members | 'Members' | No |
| `visitorLabel` | String | Custom label for visitors | 'Visitors' | No |

## Compatibility

![SPFx 1.21.1](https://img.shields.io/badge/SPFx-1.21.1-green.svg)  
![Node.js v18-22](https://img.shields.io/badge/Node.js-v18--22-green.svg)  
![SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)  
![Teams](https://img.shields.io/badge/Microsoft%20Teams-Compatible-green.svg)

## Prerequisites

- SharePoint Online tenant
- Microsoft 365 account with access to Microsoft Graph
- Node.js v18
- SharePoint Framework 1.20.2

## Installation

1. Clone the repository
2. Navigate to the project directory
3. Run `npm install`
4. Run `gulp serve` for local testing
5. Use `gulp bundle --ship` and `gulp package-solution --ship` for production deployment

## Usage Example

```typescript
// Web part configuration example
{
  roles: ['owner', 'admin', 'member'],
  itemsPerPage: 15,
  sortField: 'jobTitle',
  showSearchBox: true,
  ownerLabel: 'Site Owners',
  adminLabel: 'Administrators'
}
```

## Security and Permissions

### Required Microsoft Graph API Permissions

The following Microsoft Graph API permissions are required for full functionality:

| Permission Scope | Purpose | Type |
|-----------------|---------|------|
| `User.Read.All` | Read comprehensive user profiles | Application |
| `User.ReadBasic.All` | Read basic user profile information | Application |
| `Group.Read.All` | Read Microsoft 365 group details | Application |
| `GroupMember.Read.All` | Read group membership information | Application |
| `Presence.Read.All` | Read user presence status in Microsoft Teams | Application |
| `Sites.Read.All` | Read SharePoint site information and permissions | Application |
| `Sites.ReadWrite.All` | Read and write SharePoint site permissions | Application |

### Permission Request Process

1. Deploy the web part solution
2. Navigate to the SharePoint Admin Center
3. Go to "Advanced" > "API Access"
4. Approve the requested Microsoft Graph API permissions

### Security Considerations

- Users will only see group members they have permission to view
- Respects existing SharePoint and Microsoft 365 group access controls
- Permissions are scoped to read-only access

## Known Limitations

- Performance may vary with large group memberships
- Profile photo retrieval depends on user's Microsoft 365 profile

## Troubleshooting

- Ensure proper Microsoft Graph API permissions
- Verify network connectivity
- Check browser console for specific error messages

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

[MIT License](LICENSE)

## Disclaimer

THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED.

## Help and Support

- [Microsoft 365 Developer Community](https://aka.ms/m365dev)
- [SharePoint Framework Documentation](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft Graph Documentation](https://docs.microsoft.com/graph/overview)

## Architecture

### Key Components

- **GraphService**: Centralized Microsoft Graph API client with intelligent caching
- **ProfileImage**: Optimized image component with fallback handling
- **ErrorBoundary**: Robust error handling and recovery
- **LivePersona**: Enhanced user persona with presence integration

### Performance Features

- **Intelligent Caching**: Session-based caching with TTL
- **Lazy Loading**: Images load on demand
- **Debounced Search**: 300ms debounce for optimal performance
- **Error Recovery**: Automatic retry mechanisms
- **Memory Management**: Proper cleanup and leak prevention

## Version History

| Version | Date | Comments |
|---------|------|----------|
| 3.0.0 | February 2025 | Major refactor with fixes to retreving the right information, deeper retrival from groups such as everyone except externals etc |
| 2.0.0 | February 2025 | Major refactor with GraphService, improved caching, LivePersona integration |
| 1.0.0 | February 2025 | Initial release |