# WOP (Work Order Portal) - Setup and Testing Instructions

## Overview

The Work Order Portal (WOP) generates and persists WorkOrderIDs in the format `WO-YYYYMMDD-HHMMSS` (UTC) and provides a UI for managing work orders with sheet integration.

## Files

- **index_wop.html** - Main portal page with work order list and creation form
- **right_portal.html** - Work order details panel with Assign Engineer and Schedule actions
- **Code_wop.gs** - Apps Script backend providing Web App API
- **public/js/wop_client_create_with_generator.js** - Client-side helper for WorkOrderID generation

## Setup Instructions

### 1. Deploy Apps Script as Web App

1. Open Google Apps Script editor
2. Create a new project or use existing
3. Copy the contents of `Code_wop.gs` into the script editor
4. Update the `SPREADSHEET_ID` constant with your Google Sheets ID
5. Deploy as Web App:
   - Click **Deploy** > **New deployment**
   - Select **Web app** as deployment type
   - Set "Execute as" to **Me**
   - Set "Who has access" to **Anyone** (or as needed)
   - Click **Deploy**
6. Copy the Web App URL

### 2. Update Client Configuration

1. Open `index_wop.html` and update any APPS_SCRIPT_URL references
2. Open `public/js/wop_client_create_with_generator.js`
3. Replace `'YOUR_WEB_APP_URL_HERE'` with your Web App URL

### 3. Google Sheets Setup

The Apps Script will automatically create the required sheets and columns:

- **WorkOrders** sheet with columns:
  - WorkOrderID
  - WorkOrderText
  - EngineerID
  - EngineerName
  - ScheduledDate
  - CreatedUTC
  - LastUpdate

- **Engineers** sheet with columns:
  - EngineerID
  - EngineerName
  - EngineerEmail
  - Active

**Optional:** Pre-populate the Engineers sheet with sample data:

| EngineerID | EngineerName | EngineerEmail | Active |
|------------|--------------|---------------|--------|
| ENG001 | John Smith | john@example.com | Y |
| ENG002 | Jane Doe | jane@example.com | Y |

## Testing Instructions

### Test 1: WorkOrderID Generation Format

1. Open `index_wop.html` in a browser
2. Open browser console (F12)
3. Run: `window.generateWorkOrderId()`
4. **Expected:** Returns WorkOrderID in format `WO-YYYYMMDD-HHMMSS`
   - Example: `WO-20251207-184530`
5. Run multiple times and verify format is consistent
6. **Verify:** Format is UTC timezone (compare with current UTC time)

### Test 2: Create Work Order from UI

1. In the WOP UI, enter a description in "Work Order Description" field
2. Click "Create Work Order" button
3. **Expected:** 
   - Success message shows: "Work order created: WO-YYYYMMDD-HHMMSS"
   - New work order appears in the work orders list
4. Check Google Sheets WorkOrders tab
5. **Verify:**
   - New row added with correct WorkOrderID format
   - WorkOrderText contains entered description

### Test 3: Create Work Order Without Description

1. Leave "Work Order Description" field empty
2. Click "Create Work Order" button
3. **Expected:**
   - Work order created successfully
   - WorkOrderText defaults to ' FROM WOP'
4. Check Google Sheets
5. **Verify:** WorkOrderText column shows ' FROM WOP'

### Test 4: Assign Engineer

1. Select a work order from the list
2. Click "Assign Engineer" button
3. Select an engineer from dropdown
4. Click "Assign" button
5. **Expected:**
   - Success message: "Engineer assigned successfully"
   - Work order details update to show assigned engineer
   - Work order list item shows engineer name
6. Check Google Sheets WorkOrders tab
7. **Verify:**
   - EngineerID and EngineerName columns updated for that WorkOrderID
   - LastUpdate timestamp updated

### Test 5: Schedule Work Order

1. Select a work order from the list
2. Click "Schedule" button
3. Select a date
4. Click "Schedule" button
5. **Expected:**
   - Success message: "Work order scheduled successfully"
   - Work order details update to show scheduled date
6. Check Google Sheets WorkOrders tab
7. **Verify:**
   - ScheduledDate column updated for that WorkOrderID
   - LastUpdate timestamp updated

### Test 6: Client Helper Function

1. Open browser console
2. Run: `window.WOPClient.validateWorkOrderIdFormat('WO-20251207-184530')`
3. **Expected:** Returns `true`
4. Run: `window.WOPClient.validateWorkOrderIdFormat('WO1011')`
5. **Expected:** Returns `false`
6. Run: `window.WOPClient.extractDateFromWorkOrderId('WO-20251207-184530')`
7. **Expected:** Returns Date object for 2025-12-07 18:45:30 UTC

### Test 7: UI Verification

1. **Appbar:**
   - Title: "OBSI WORK ORDER PROTAL" (center-justified)
   - Left side: "© 2025 OBSI"
   - Right side: Build label
2. **No debug panel visible**
3. **Action buttons:**
   - Styled with .btn and .btn-ghost classes
   - Consistent sizing

### Test 8: Fallback Behavior

1. Open browser console
2. Run: `window.createWorkOrderWithGenerator({ workOrderText: 'Test fallback' })`
3. **Expected:**
   - WorkOrderID generated locally if server unavailable
   - Work order created successfully
4. Check format with: `window.validateWorkOrderIdFormat(result.WorkOrderID)`
5. **Expected:** Returns `true`

## API Endpoints

### GET Endpoints

- `?op=getEngineers` - Returns list of engineers

### POST Endpoints (JSON body)

- `{ action: 'assignEngineer', workOrderId, engineerId, engineerName }` - Assigns engineer to work order
- `{ action: 'scheduleWorkOrder', workOrderId, scheduleDate }` - Sets scheduled date
- `{ action: 'createRow', ...workOrderData }` - Creates new work order row

## Notes

- **No ASP files modified:** This PR only adds WOP-specific files
- **UTC timestamps:** All WorkOrderIDs use UTC timezone
- **No note fields:** Assign Engineer and Schedule dialogs do not include note fields per requirements
- **Default text:** WorkOrderText defaults to ' FROM WOP' when missing
- **Format validation:** Use `validateWorkOrderIdFormat()` to verify WorkOrderID format
- **Sheet auto-creation:** Missing columns/sheets are automatically created on first use

## Troubleshooting

### WorkOrderID not generated
- Check browser console for errors
- Verify `window.generateWorkOrderId()` function is available
- Fallback to local generator if server unavailable

### Engineers not loading
- Verify Engineers sheet exists in Google Sheets
- Check APPS_SCRIPT_URL is correctly configured
- Add sample engineers to sheet manually

### API calls failing
- Verify Apps Script Web App is deployed
- Check "Execute as: Me" and appropriate access permissions
- Verify APPS_SCRIPT_URL is correct in client files

## Browser Compatibility

- Chrome/Edge: Fully supported
- Firefox: Fully supported
- Safari: Fully supported
- Mobile browsers: Responsive design included
