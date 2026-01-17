# Shift Manager - Project Guide for Claude

## Project Overview

**Shift Manager** is a client-side web application for employee scheduling and shift management. Built with React, it helps organizations manage employee records, create work schedules, track shift assignments, and ensure fair distribution of shifts through a "justice points" system.

**Key Characteristics:**
- Single-page React application (monolithic component structure)
- Client-side only - no backend, all data stored in localStorage
- Bilingual support (English & Hebrew with RTL)
- Excel import/export capabilities
- Justice points system for fair shift distribution

## Technology Stack

**Frontend:**
- React 19.2.3 with React DOM
- Create React App (react-scripts 5.0.1)
- Tailwind CSS 4.1.18 for styling
- Lucide React for icons

**Data & Files:**
- XLSX library for Excel file handling
- localStorage for data persistence

**Testing:**
- @testing-library/react, jest-dom, user-event

## Project Structure

```
shift-manager/
├── public/                  # Static assets
│   ├── index.html          # Entry HTML with Tailwind CDN
│   ├── manifest.json       # PWA configuration
│   └── logo files
├── src/
│   ├── App.js              # Main application (1500+ lines, all logic)
│   ├── App.css             # Legacy CSS (minimal)
│   ├── App.test.js         # Unit tests
│   ├── index.js            # React entry point
│   ├── index.css           # Global styles
│   ├── tailwind.config.js  # Tailwind configuration
│   └── postcss.config.js   # PostCSS configuration
└── package.json
```

## Architecture Patterns

### State Management
- Uses React Hooks (useState, useEffect, useMemo, useRef)
- All state managed in App.js
- Debounced auto-save to localStorage (1 second delay)
- No external state management library

### Data Model

**Employee:**
```javascript
{
  id: timestamp,
  name: string,
  personalNumber: string,
  sex: 'Male' | 'Female' | 'Other',
  status: string,  // Custom status values
  department: string
}
```

**Shift:**
```javascript
{
  id: timestamp,
  employeeId: number | null,  // null = unassigned
  date: string,  // YYYY-MM-DD
  startTime: string,  // datetime-local format
  endTime: string,    // datetime-local format
  role: string,  // חול, שבת, שישי, חג
  repeatType: 'none' | 'daily' | 'weekly' | 'monthly',
  repeatUntil: string  // Optional end date
}
```

**Shift Types with Point Values:**
- חול (weekday): 1 point
- שישי (Friday): 1.5 points
- שבת (Saturday): 2 points
- חג (holiday): 3 points

### Key Features

**Employee Management:**
- CRUD operations for employees
- Excel import with template download
- Search/filter by name, ID, department, status
- Custom status management

**Shift Scheduling:**
- CRUD operations for shifts
- Assign/unassign employees
- Repeating shifts (daily/weekly/monthly)
- Conflict detection (overlapping shifts)
- Shift swap functionality
- Auto-distribution of unassigned shifts
- Multiple calendar views (List, Week, Month)

**Reporting:**
- Employee statistics with shift counts
- Justice points calculation for fairness
- Export to Excel
- Print schedules

**UI/UX:**
- Dark mode toggle
- Bilingual (English/Hebrew with RTL)
- Responsive design
- Real-time filtering
- Save status indicators
- Confirmation dialogs

## Important Implementation Notes

### Styling Approach
- All styles defined as inline JavaScript objects in App.js
- Tailwind CSS used via CDN link in public/index.html
- Minimal use of separate CSS files (App.css, index.css are legacy)

### Data Persistence
- Everything saved to localStorage
- Keys: 'employees', 'shifts', 'language', 'isDarkMode'
- Auto-save with debouncing to prevent excessive writes
- No server API calls

### Date/Time Handling
- Uses native HTML datetime-local inputs
- JavaScript Date objects for calculations
- Date format: YYYY-MM-DD for storage
- Conflict detection based on date ranges

### Localization
- Translations object embedded in App.js
- Language toggle persists to localStorage
- RTL support for Hebrew

### Justice Points System
- Fairness metric calculated per employee
- Based on shift types and point values
- Used to identify uneven shift distribution
- Displayed in Reports tab

## Common Development Tasks

### Adding a New Feature
1. Most logic is in `src/App.js` - this is the primary file to modify
2. Add new state with `useState` if needed
3. Add UI elements within the appropriate tab section
4. Update localStorage persistence if new data types added
5. Add translations to both 'en' and 'he' in translations object

### Modifying Shift Types
1. Update `shiftTypeInfo` object in App.js
2. Add points, colors, labels (Hebrew & English)
3. Update UI components that render shift types

### Adding New Reports
1. Add logic to Reports tab section in App.js
2. Use `useMemo` for calculated statistics
3. Consider Excel export functionality

### Styling Changes
1. Modify inline style objects in App.js
2. Or add Tailwind classes to className attributes
3. Dark mode: use ternary operators based on `isDarkMode` state

### Testing
- Run tests with `npm test`
- Tests located in `src/App.test.js`
- Uses @testing-library/react

## Development Commands

```bash
npm start          # Run development server (localhost:3000)
npm build          # Create production build
npm test           # Run tests in watch mode
npm eject          # Eject from Create React App (irreversible)
```

## Known Technical Debt

1. **Monolithic Component**: All logic in single 1500+ line App.js file
   - Consider refactoring into smaller components
   - Separate concerns (employees, shifts, reports into modules)

2. **No Backend**: All data in localStorage
   - Consider adding database backend for multi-user support
   - No data backup or sync across devices

3. **Inline Styles**: Makes maintenance harder
   - Consider moving to CSS modules or styled-components
   - Or fully embrace Tailwind utility classes

4. **No State Management Library**: Complex state in single component
   - Consider Redux, Zustand, or Context API for scaling

5. **No TypeScript**: Potential for type-related bugs
   - Consider migration to TypeScript for better type safety

## Security & Privacy

- All data stored client-side in browser localStorage
- No server communication or external API calls
- Data persists only in user's browser
- Clearing browser data will delete all information
- No authentication or authorization system

## Browser Compatibility

- Modern browsers with ES6+ support required
- localStorage API required
- File API for Excel import/export
- datetime-local input support needed

## Future Enhancement Ideas

- Backend API with database
- User authentication and multi-tenant support
- Real-time collaboration
- Mobile app version
- Advanced reporting with charts
- Notifications for upcoming shifts
- Employee availability preferences
- Shift trading/swap requests
- Calendar integrations (Google Calendar, Outlook)
- Audit log for changes
- Backup/restore functionality
- Advanced auto-scheduling algorithms

## Troubleshooting

**Data Loss Issues:**
- Check browser localStorage limits (typically 5-10MB)
- Verify localStorage is enabled in browser settings
- Check for browser extensions blocking localStorage

**Excel Import Not Working:**
- Ensure file is .xlsx or .xls format
- Check template format matches expected structure
- Verify XLSX library is installed

**Conflicts Not Detected:**
- Verify date/time formats are correct
- Check conflict detection logic in `hasConflict()` function
- Ensure shifts have valid start/end times

**Language/RTL Issues:**
- Check language state in localStorage
- Verify translations object has both 'en' and 'he' keys
- CSS direction is managed by `dir` attribute on elements

## Contributing Guidelines

When making changes:
1. Test in both English and Hebrew modes
2. Test with dark mode enabled
3. Verify localStorage persistence works
4. Check for shift conflicts with new logic
5. Ensure Excel import/export still functions
6. Test with both assigned and unassigned shifts
7. Verify justice points calculations remain accurate
8. Check responsive design on mobile viewports

## Contact & Resources

- Repository: (check git remote for URL)
- Main branch: master
- Current worktree: unruffled-tharp
- Issues: Report bugs via GitHub issues
- Documentation: README.md for user-facing docs

---

**Last Updated:** 2026-01-17
**Primary File to Modify:** src/App.js (contains ~95% of application logic)
