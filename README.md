# Batt Fix ⚡

A mobile-first battery repair and technician attendance app designed for GitHub + Vercel deployment.

## What is included

- Battery Repair: Gen 4 / Gen 5 response flow.
- Real-time Punch In and Punch Out.
- Attendance timestamps are generated on the Vercel server, not entered by the technician.
- Duplicate Punch-In / Punch-Out protection.
- Punch-Out must be later than Punch-In.
- Completed attendance is locked by the API design: only an open record can be punched out.
- Technician ID and name captured with each attendance record.
- Optional location and device metadata.
- Attendance history with search, date filters and status filters.
- CSV export.
- Dashboard and local offline cache.
- iPhone/Safari-friendly responsive UI and safe-area handling.
- Share/download repair response images.
- SharePoint List integration through Microsoft Graph, with secrets kept server-side.

## Deploy to Vercel

1. Create a GitHub repository and upload this entire folder.
2. Import the repository into Vercel.
3. Vercel will use `index.html` as the frontend and `api/attendance.js` as the serverless API.
4. Add these Vercel Environment Variables:

```text
SP_TENANT_ID=your-entra-tenant-id
SP_CLIENT_ID=your-entra-app-client-id
SP_CLIENT_SECRET=your-entra-app-client-secret
SP_SITE_ID=your-sharepoint-site-id
SP_LIST_ID=your-sharepoint-list-id
```

5. Redeploy.

## SharePoint List columns

Create a SharePoint List with these columns. Use the exact internal names where possible:

- Title: Single line of text
- TechnicianID: Single line of text
- TechnicianName: Single line of text
- PunchIn: Date and Time
- PunchOut: Date and Time
- WorkingMinutes: Number
- Status: Single line of text
- Location: Multiple lines of text or Single line of text
- DeviceInfo: Multiple lines of text or Single line of text
- BatteriesRepaired: Number
- CasesReplaced: Number
- RecordHash: Single line of text

The app uses Microsoft Graph application credentials from Vercel only. Never put the client secret in the frontend.

## Microsoft Entra / Graph permission

Register an Entra ID application and grant it the minimum SharePoint/Graph application permissions needed for the target site/list. Prefer site-scoped access where your tenant supports it. Grant admin consent, then copy the tenant ID, client ID and secret into Vercel.

## Important production note

The browser cannot make a trustworthy claim about its own clock. This version therefore never accepts a Punch-In or Punch-Out timestamp from the client. The API creates both timestamps from the server clock. Client-side validation is only for user experience; the server remains the authority.

For stronger identity assurance, put the app behind Microsoft Entra authentication or another authenticated gateway. Technician ID + name alone should not be treated as a high-assurance identity mechanism.
