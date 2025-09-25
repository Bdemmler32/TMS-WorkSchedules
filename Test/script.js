// Global variables
let employeeData = {};
let currentWeekStart = null;
let currentWeekType = 1; // 1 or 2
let selectedEmployees = []; // For filtering
let isFilterActive = false;
let nameSortOrder = 'first'; // 'first' or 'last'
let officeHoursOnly = false;
let calendarDate = new Date(); // Current month/year being displayed in calendar
let selectedWeekStart = null; // Week selected in calendar
let lastUpdatedDate = null; // Last updated date from Admin sheet

// Week 1 starts on September 13, 2025
const WEEK_1_START = new Date('2025-09-13T00:00:00');

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    loadScheduleData();
    
    // Start collapsed by default - add null checks
    const unifiedContainer = document.getElementById('unifiedContainer');
    const expandableContent = document.getElementById('expandableContent');
    
    if (unifiedContainer) unifiedContainer.classList.add('collapsed');
    if (expandableContent) expandableContent.classList.add('collapsed');
});

function initializeEventListeners() {
    // Navigation buttons
    document.getElementById('prevWeek').addEventListener('click', () => navigateWeek(-1));
    document.getElementById('nextWeek').addEventListener('click', () => navigateWeek(1));
    
    // Jump to Today button
    document.getElementById('jumpTodayBtn').addEventListener('click', jumpToToday);
    
    // Filter button
    document.getElementById('filterBtn').addEventListener('click', handleFilterBtn);
    
    // Expand/Collapse text button
    document.getElementById('expandCollapseText').addEventListener('click', toggleScheduleCollapse);
    
    // Date range click to open calendar
    document.getElementById('dateRange').addEventListener('click', openCalendarModal);
    
    // Modal events - Employee modal
    document.getElementById('closeModal').addEventListener('click', closeModal);
    document.getElementById('employeeModal').addEventListener('click', (e) => {
        if (e.target.id === 'employeeModal') closeModal();
    });
    
    // Filter modal events - NO CLICK OUTSIDE TO CLOSE
    document.getElementById('closeFilterModal').addEventListener('click', closeFilterModal);
    
    // Calendar modal events
    document.getElementById('closeCalendarModal').addEventListener('click', closeCalendarModal);
    document.getElementById('calendarModal').addEventListener('click', (e) => {
        if (e.target.id === 'calendarModal') closeCalendarModal();
    });
    
    // Calendar navigation
    document.getElementById('prevMonth').addEventListener('click', () => navigateCalendarMonth(-1));
    document.getElementById('nextMonth').addEventListener('click', () => navigateCalendarMonth(1));
    
    // Calendar actions
    document.getElementById('todayBtn').addEventListener('click', selectTodayInCalendar);
    document.getElementById('cancelCalendarBtn').addEventListener('click', closeCalendarModal);
    document.getElementById('applyCalendarBtn').addEventListener('click', applyCalendarSelection);
    
    // Filter actions
    document.getElementById('selectAllBtn').addEventListener('click', selectAllEmployees);
    document.getElementById('deselectAllBtn').addEventListener('click', deselectAllEmployees);
    document.getElementById('applyFilterBtn').addEventListener('click', applyEmployeeFilter);
    document.getElementById('officeHoursCheckbox').addEventListener('change', toggleOfficeHoursFilter);
    
    // Employee search
    document.getElementById('employeeSearch').addEventListener('input', filterEmployeeList);
}

async function loadScheduleData() {
    try {
        showLoading(true);
        
        const response = await fetch('TMS-WorkSchedules.xlsx');
        if (!response.ok) {
            throw new Error('Could not load Excel file');
        }
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        processWorkbookData(workbook);
        
        // Set current week based on today's date
        setCurrentWeek();
        updateDisplay();
        
        showLoading(false);
        
    } catch (error) {
        console.error('Error loading schedule:', error);
        showError();
    }
}

function processWorkbookData(workbook) {
    const validSheets = workbook.SheetNames.filter(name => 
        name !== 'NewEmployee' && name !== 'FormTools' && name !== 'Admin'
    );
    
    // Read last updated date from Admin sheet if it exists
    if (workbook.SheetNames.includes('Admin')) {
        const adminSheet = workbook.Sheets['Admin'];
        const lastUpdatedCell = adminSheet['E10'];
        if (lastUpdatedCell && lastUpdatedCell.v) {
            lastUpdatedDate = lastUpdatedCell.v;
        }
    }
    
    const daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    const dayColumns = {
        'Monday': ['I', 'J', 'K'],
        'Tuesday': ['L', 'M', 'N'], 
        'Wednesday': ['O', 'P', 'Q'],
        'Thursday': ['R', 'S', 'T'],
        'Friday': ['U', 'V', 'W']
    };
    
    validSheets.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const employeeName = sheet['C1']?.v || sheetName;
        
        employeeData[employeeName] = {
            sheetName: sheetName,
            week1: {},
            week2: {}
        };
        
        // Process Week 1 (rows 9-13)
        for (let row = 9; row <= 13; row++) {
            processDayData(sheet, row, daysOfWeek, dayColumns, employeeName, 'week1');
        }
        
        // Process Week 2 (rows 24-28)
        for (let row = 24; row <= 28; row++) {
            processDayData(sheet, row, daysOfWeek, dayColumns, employeeName, 'week2');
        }
    });
}

function processDayData(sheet, row, daysOfWeek, dayColumns, employeeName, weekType) {
    daysOfWeek.forEach(day => {
        const [startCol, endCol, locationCol] = dayColumns[day];
        const startTime = sheet[`${startCol}${row}`]?.v;
        const endTime = sheet[`${endCol}${row}`]?.v;
        const location = sheet[`${locationCol}${row}`]?.v;
        
        if (startTime && endTime && location) {
            if (!employeeData[employeeName][weekType][day]) {
                employeeData[employeeName][weekType][day] = [];
            }
            
            const startTimeFormatted = excelDateToTime(startTime);
            const endTimeFormatted = excelDateToTime(endTime);
            
            // Ensure start time comes before end time
            const [finalStart, finalEnd] = orderTimes(startTimeFormatted, endTimeFormatted);
            
            employeeData[employeeName][weekType][day].push({
                startTime: finalStart,
                endTime: finalEnd,
                location: location.trim(),
                block: weekType === 'week1' ? row - 8 : row - 23
            });
        }
    });
}

function excelDateToTime(excelDate) {
    if (!excelDate) return null;
    
    // Handle both date objects and decimal time values
    let timeValue;
    if (excelDate instanceof Date) {
        timeValue = excelDate.getHours() + (excelDate.getMinutes() / 60);
    } else {
        // Excel decimal time (e.g., 0.5 = 12:00 PM)
        timeValue = excelDate * 24;
    }
    
    const hours = Math.floor(timeValue);
    const minutes = Math.round((timeValue - hours) * 60);
    
    // Convert to 12-hour format
    const period = hours >= 12 ? 'PM' : 'AM';
    const displayHours = hours === 0 ? 12 : (hours > 12 ? hours - 12 : hours);
    
    return `${displayHours}:${minutes.toString().padStart(2, '0')} ${period}`;
}

function orderTimes(startTime, endTime) {
    // Don't reorder times - keep them as they appear in the Excel file
    return [startTime, endTime];
}

function timeToMinutes(timeStr) {
    const [time, period] = timeStr.split(' ');
    const [hours, minutes] = time.split(':').map(Number);
    let totalMinutes = (hours % 12) * 60 + minutes;
    if (period === 'PM' && hours !== 12) totalMinutes += 720;
    if (period === 'AM' && hours === 12) totalMinutes = minutes;
    return totalMinutes;
}

function setCurrentWeek() {
    const today = new Date();
    const daysSinceWeek1 = Math.floor((today - WEEK_1_START) / (1000 * 60 * 60 * 24));
    const weekNumber = Math.floor(daysSinceWeek1 / 7);
    
    currentWeekType = (weekNumber % 2) + 1;
    
    // Set current week start to the beginning of this week period
    const weeksFromStart = Math.floor(daysSinceWeek1 / 7);
    currentWeekStart = new Date(WEEK_1_START);
    currentWeekStart.setDate(currentWeekStart.getDate() + (weeksFromStart * 7));
    
    // Update week display
    updateWeekDisplay();
}

function navigateWeek(direction) {
    currentWeekStart.setDate(currentWeekStart.getDate() + (direction * 7));
    
    // Determine week type based on weeks from original start
    const daysSinceStart = Math.floor((currentWeekStart - WEEK_1_START) / (1000 * 60 * 60 * 24));
    const weekNumber = Math.floor(daysSinceStart / 7);
    currentWeekType = (weekNumber % 2) + 1;
    
    updateWeekDisplay();
    updateDisplay();
}

function updateDisplay() {
    updateDateRange();
    updateWeekDisplay();
    renderScheduleGrid();
}

function updateWeekDisplay() {
    const weekDisplay = document.getElementById('weekDisplay');
    if (weekDisplay) {
        if (currentWeekType === 1) {
            weekDisplay.textContent = 'Week 1 - Pay Week';
        } else {
            weekDisplay.textContent = 'Week 2';
        }
    }
}

function updateDateRange() {
    const weekEnd = new Date(currentWeekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    const startStr = currentWeekStart.toLocaleDateString('en-US', options);
    const endStr = weekEnd.toLocaleDateString('en-US', options);
    
    const dateRange = document.getElementById('dateRange');
    if (dateRange) {
        dateRange.textContent = `${startStr} - ${endStr}`;
    }
}

// Calendar Functions
function openCalendarModal() {
    const modal = document.getElementById('calendarModal');
    
    // Set calendar to current week's month
    calendarDate = new Date(currentWeekStart);
    selectedWeekStart = new Date(currentWeekStart);
    
    renderCalendar();
    modal.style.display = 'flex';
}

function closeCalendarModal() {
    document.getElementById('calendarModal').style.display = 'none';
}

function navigateCalendarMonth(direction) {
    calendarDate.setMonth(calendarDate.getMonth() + direction);
    renderCalendar();
}

function selectTodayInCalendar() {
    const today = new Date();
    selectedWeekStart = getWeekStartFromDate(today);
    calendarDate = new Date(today);
    renderCalendar();
}

function applyCalendarSelection() {
    if (selectedWeekStart) {
        currentWeekStart = new Date(selectedWeekStart);
        
        // Calculate week type
        const daysSinceStart = Math.floor((currentWeekStart - WEEK_1_START) / (1000 * 60 * 60 * 24));
        const weekNumber = Math.floor(daysSinceStart / 7);
        currentWeekType = (weekNumber % 2) + 1;
        
        updateDisplay();
        closeCalendarModal();
    }
}

function getWeekStartFromDate(date) {
    // Calculate which work week this date belongs to
    const daysSinceWeek1 = Math.floor((date - WEEK_1_START) / (1000 * 60 * 60 * 24));
    const weekNumber = Math.floor(daysSinceWeek1 / 7);
    
    const weekStart = new Date(WEEK_1_START);
    weekStart.setDate(weekStart.getDate() + (weekNumber * 7));
    return weekStart;
}

function renderCalendar() {
    const grid = document.getElementById('calendarGrid');
    const monthYear = document.getElementById('calendarMonthYear');
    
    const year = calendarDate.getFullYear();
    const month = calendarDate.getMonth();
    
    // Update month/year display
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    monthYear.textContent = `${monthNames[month]} ${year}`;
    
    // Clear grid
    grid.innerHTML = '';
    
    // Add day headers
    const dayHeaders = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    dayHeaders.forEach(day => {
        const header = document.createElement('div');
        header.className = 'calendar-day-header';
        header.textContent = day;
        grid.appendChild(header);
    });
    
    // Get first day of month and number of days
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();
    const startingDayOfWeek = firstDay.getDay();
    
    // Add empty cells for days before month starts
    for (let i = 0; i < startingDayOfWeek; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'calendar-day empty';
        grid.appendChild(emptyCell);
    }
    
    // Add days of the month
    for (let day = 1; day <= daysInMonth; day++) {
        const dayCell = document.createElement('div');
        dayCell.className = 'calendar-day';
        dayCell.textContent = day;
        
        const cellDate = new Date(year, month, day);
        const weekStart = getWeekStartFromDate(cellDate);
        
        // Check if this day is in the selected week
        if (selectedWeekStart) {
            const weekEnd = new Date(selectedWeekStart);
            weekEnd.setDate(weekEnd.getDate() + 6);
            
            if (cellDate >= selectedWeekStart && cellDate <= weekEnd) {
                dayCell.classList.add('selected-week');
            }
        }
        
        // Check if this is today
        const today = new Date();
        if (cellDate.toDateString() === today.toDateString()) {
            dayCell.classList.add('today');
        }
        
        // Add click handler
        dayCell.addEventListener('click', () => {
            selectedWeekStart = weekStart;
            renderCalendar(); // Re-render to show selection
        });
        
        grid.appendChild(dayCell);
    }
}

function sortEmployees(employees) {
    return employees.sort((a, b) => {
        if (nameSortOrder === 'first') {
            return a.localeCompare(b);
        } else {
            // Sort by last name
            const lastNameA = a.split(' ').pop();
            const lastNameB = b.split(' ').pop();
            return lastNameA.localeCompare(lastNameB);
        }
    });
}

function toggleNameSort() {
    nameSortOrder = nameSortOrder === 'first' ? 'last' : 'first';
    updateDisplay();
}

function toggleOfficeHoursFilter() {
    const checkbox = document.getElementById('officeHoursCheckbox');
    officeHoursOnly = checkbox.checked;
}

function hasOfficeHours(employee, weekData) {
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    return days.some(day => {
        const daySchedule = weekData[day] || [];
        return daySchedule.some(block => block.location.toLowerCase() === 'office');
    });
}

function filterOfficeHoursOnly(weekData) {
    const filteredData = {};
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    
    days.forEach(day => {
        const daySchedule = weekData[day] || [];
        const officeBlocks = daySchedule.filter(block => 
            block.location.toLowerCase() === 'office'
        );
        if (officeBlocks.length > 0) {
            filteredData[day] = officeBlocks;
        }
    });
    
    return filteredData;
}

function getCurrentDayOfWeek() {
    const today = new Date();
    const dayIndex = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    
    // Convert to our format (Monday = 0, Tuesday = 1, etc.)
    // Return -1 if today is weekend (not shown in schedule)
    if (dayIndex === 0 || dayIndex === 6) return -1; // Sunday or Saturday
    return dayIndex - 1; // Monday = 0, Tuesday = 1, etc.
}

function isCurrentWeek() {
    const today = new Date();
    const todayStart = new Date(today);
    todayStart.setHours(0, 0, 0, 0);
    
    const currentWeekEnd = new Date(currentWeekStart);
    currentWeekEnd.setDate(currentWeekEnd.getDate() + 6);
    currentWeekEnd.setHours(23, 59, 59, 999);
    
    return todayStart >= currentWeekStart && todayStart <= currentWeekEnd;
}

function renderScheduleGrid() {
    const scheduleGrid = document.getElementById('scheduleGrid');
    if (!scheduleGrid) return;
    
    scheduleGrid.innerHTML = '';
    
    let employees = Object.keys(employeeData);
    employees = sortEmployees(employees);
    
    const colors = ['color-1', 'color-2', 'color-3', 'color-4', 'color-5', 'color-6', 
                   'color-7', 'color-8', 'color-9', 'color-10', 'color-11', 'color-12'];
    
    // Filter employees based on selection and schedule data
    let filteredEmployees;
    if (isFilterActive && selectedEmployees.length > 0) {
        filteredEmployees = selectedEmployees.filter(employeeName => {
            const employee = employeeData[employeeName];
            if (!employee) return false;
            let weekData = currentWeekType === 1 ? employee.week1 : employee.week2;
            
            // Apply office filter if active
            if (officeHoursOnly) {
                if (!hasOfficeHours(employee, weekData)) return false;
                weekData = filterOfficeHoursOnly(weekData);
            }
            
            return hasScheduleData(weekData);
        });
        filteredEmployees = sortEmployees(filteredEmployees);
    } else {
        filteredEmployees = employees.filter(employeeName => {
            const employee = employeeData[employeeName];
            let weekData = currentWeekType === 1 ? employee.week1 : employee.week2;
            
            // Apply office filter if active
            if (officeHoursOnly && !hasOfficeHours(employee, weekData)) {
                return false;
            }
            
            return hasScheduleData(weekData);
        });
    }
    
    // Set grid rows: 1 header row + number of filtered employees
    scheduleGrid.style.gridTemplateRows = `auto repeat(${filteredEmployees.length}, auto)`;
    
    // Get current day info
    const currentDayIndex = getCurrentDayOfWeek();
    const isThisWeek = isCurrentWeek();
    
    // Create header cells
    const headers = ['Employee', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    headers.forEach((header, index) => {
        const headerCell = document.createElement('div');
        let className = `grid-cell header-cell ${index === 0 ? 'employee-header' : ''}`;
        
        // Highlight current day if this is the current week
        if (isThisWeek && index > 0 && index - 1 === currentDayIndex) {
            className += ' current-day';
        }
        
        headerCell.className = className;
        
        if (index === 0) {
            // Employee header with sort button
            headerCell.innerHTML = `
                <div class="employee-header-content">
                    <span>Employee</span>
                    <button class="name-sort-btn ${nameSortOrder === 'last' ? 'active' : ''}" id="nameSortBtn" title="Toggle First/Last Name Sort">
                        <i class="fas fa-sort-alpha-${nameSortOrder === 'first' ? 'down' : 'up'}"></i>
                    </button>
                </div>
            `;
            
            // Add click event to sort button
            setTimeout(() => {
                const sortBtn = document.getElementById('nameSortBtn');
                if (sortBtn) {
                    sortBtn.addEventListener('click', toggleNameSort);
                }
            }, 0);
        } else {
            headerCell.textContent = header;
        }
        
        scheduleGrid.appendChild(headerCell);
    });
    
    // Create employee rows
    filteredEmployees.forEach((employeeName, employeeIndex) => {
        const employee = employeeData[employeeName];
        let weekData = currentWeekType === 1 ? employee.week1 : employee.week2;
        
        // Apply office filter to week data if active
        if (officeHoursOnly) {
            weekData = filterOfficeHoursOnly(weekData);
        }
        
        const originalIndex = employees.indexOf(employeeName);
        const colorClass = colors[originalIndex % colors.length]; // Use original index for consistent colors
        
        // Store row cells for hover effect
        const rowCells = [];
        
        // Employee info cell
        const employeeCell = document.createElement('div');
        employeeCell.className = 'grid-cell employee-cell';
        employeeCell.onclick = () => openEmployeeModal(employeeName);
        
        const initials = getInitials(employeeName);
        employeeCell.innerHTML = `
            <div class="employee-initials ${colorClass}">${initials}</div>
            <div class="employee-name">${formatEmployeeName(employeeName)}</div>
        `;
        
        scheduleGrid.appendChild(employeeCell);
        rowCells.push(employeeCell);
        
        // Day cells
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        days.forEach((day, dayIndex) => {
            const dayCell = document.createElement('div');
            let className = 'grid-cell day-cell';
            
            // Highlight current day if this is the current week
            if (isThisWeek && dayIndex === currentDayIndex) {
                className += ' current-day';
            }
            
            dayCell.className = className;
            
            const daySchedule = weekData[day] || [];
            daySchedule.forEach(block => {
                const workBlock = document.createElement('div');
                workBlock.className = `work-block ${block.location.toLowerCase() === 'office' ? 'office' : 'remote'}`;
                
                const icon = block.location.toLowerCase() === 'office' ? '<i class="fas fa-building"></i>' : '<i class="fas fa-home"></i>';
                workBlock.innerHTML = `
                    <span class="work-block-icon">${icon}</span>
                    <span class="work-block-time">${block.startTime} - ${block.endTime}</span>
                `;
                
                dayCell.appendChild(workBlock);
            });
            
            scheduleGrid.appendChild(dayCell);
            rowCells.push(dayCell);
        });
        
        // Add hover effects to entire employee row
        rowCells.forEach(cell => {
            cell.addEventListener('mouseenter', () => {
                rowCells.forEach(rowCell => rowCell.classList.add('employee-row-hover'));
            });
            
            cell.addEventListener('mouseleave', () => {
                rowCells.forEach(rowCell => rowCell.classList.remove('employee-row-hover'));
            });
        });
    });
    
    updateFilterResults(filteredEmployees.length, employees.length);
    updateLastUpdatedDisplay();
}

function updateLastUpdatedDisplay() {
    // Add or update the last updated date display in header
    let lastUpdatedElement = document.getElementById('lastUpdatedHeader');
    
    if (lastUpdatedDate && lastUpdatedElement) {
        // Format the date - handle both Excel serial numbers and actual values
        let dateString = lastUpdatedDate;
        
        // Check if it's an Excel serial date number (typically > 1000)
        if (typeof lastUpdatedDate === 'number' && lastUpdatedDate > 1000) {
            // Convert Excel serial number to JavaScript date with timezone correction
            const excelDate = new Date((lastUpdatedDate - 25569) * 86400 * 1000);
            // Add timezone offset to prevent date shifting
            const correctedDate = new Date(excelDate.getTime() + (excelDate.getTimezoneOffset() * 60 * 1000));
            dateString = correctedDate.toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'numeric',
                day: 'numeric'
            });
        } else if (lastUpdatedDate instanceof Date) {
            // It's already a date object
            dateString = lastUpdatedDate.toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'numeric',
                day: 'numeric'
            });
        } else if (typeof lastUpdatedDate === 'string') {
            // If it's a string, try to parse it as a date
            const parsedDate = new Date(lastUpdatedDate);
            if (!isNaN(parsedDate.getTime())) {
                dateString = parsedDate.toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'numeric',
                    day: 'numeric'
                });
            } else {
                // If it can't be parsed as a date, use it as-is
                dateString = lastUpdatedDate.toString();
            }
        } else {
            // Use the raw value converted to string
            dateString = lastUpdatedDate.toString();
        }
        
        lastUpdatedElement.textContent = `Last Updated: ${dateString}`;
        lastUpdatedElement.style.display = 'block';
    } else if (lastUpdatedElement) {
        lastUpdatedElement.style.display = 'none';
    }
}

function hasScheduleData(weekData) {
    return Object.keys(weekData).length > 0 && 
           Object.values(weekData).some(daySchedule => daySchedule && daySchedule.length > 0);
}

function getInitials(name) {
    return name.split(' ')
               .map(word => word.charAt(0).toUpperCase())
               .join('')
               .substring(0, 2);
}

function formatEmployeeName(name) {
    const parts = name.split(' ');
    if (parts.length >= 2) {
        const firstName = parts[0];
        const lastName = parts.slice(1).join(' ');
        return `<div class="employee-first-name">${firstName}</div><div class="employee-last-name">${lastName}</div>`;
    }
    return `<div class="employee-first-name">${name}</div>`;
}

function openEmployeeModal(employeeName) {
    const modal = document.getElementById('employeeModal');
    const employee = employeeData[employeeName];
    
    document.getElementById('modalEmployeeName').textContent = `${employeeName} - Schedule`;
    
    // Update week titles
    const currentWeekTitle = currentWeekType === 1 ? 'Week 1 - Pay Week (Current)' : 'Week 2 (Current)';
    const otherWeekTitle = currentWeekType === 1 ? 'Week 2' : 'Week 1 - Pay Week';
    
    document.getElementById('currentWeekTitle').textContent = currentWeekTitle;
    document.getElementById('otherWeekTitle').textContent = otherWeekTitle;
    
    // Render current week
    const currentWeekData = currentWeekType === 1 ? employee.week1 : employee.week2;
    renderModalWeek('currentWeekGrid', currentWeekData);
    
    // Render other week
    const otherWeekData = currentWeekType === 1 ? employee.week2 : employee.week1;
    renderModalWeek('otherWeekGrid', otherWeekData);
    
    modal.style.display = 'flex';
}

function renderModalWeek(gridId, weekData) {
    const grid = document.getElementById(gridId);
    grid.innerHTML = '';
    
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    
    days.forEach(day => {
        const dayDiv = document.createElement('div');
        dayDiv.className = 'modal-day';
        
        const header = document.createElement('div');
        header.className = 'modal-day-header';
        header.textContent = day;
        dayDiv.appendChild(header);
        
        const content = document.createElement('div');
        content.className = 'modal-day-content';
        
        const daySchedule = weekData[day] || [];
        
        if (daySchedule.length === 0) {
            const noWork = document.createElement('div');
            noWork.className = 'no-work';
            noWork.textContent = 'No scheduled work';
            content.appendChild(noWork);
        } else {
            daySchedule.forEach(block => {
                const workBlock = document.createElement('div');
                workBlock.className = `modal-work-block ${block.location.toLowerCase() === 'office' ? 'office' : 'remote'}`;
                
                const icon = block.location.toLowerCase() === 'office' ? '<i class="fas fa-building"></i>' : '<i class="fas fa-home"></i>';
                workBlock.innerHTML = `
                    <span class="work-block-icon">${icon}</span>
                    <span class="work-block-time">${block.startTime} - ${block.endTime}</span>
                `;
                
                content.appendChild(workBlock);
            });
        }
        
        dayDiv.appendChild(content);
        grid.appendChild(dayDiv);
    });
}

function closeModal() {
    document.getElementById('employeeModal').style.display = 'none';
}

// Filter functionality
function handleFilterBtn() {
    if (isFilterActive || officeHoursOnly) {
        clearEmployeeFilter();
    } else {
        openEmployeeFilterModal();
    }
}

function openEmployeeFilterModal() {
    const modal = document.getElementById('employeeFilterModal');
    populateEmployeeFilterList();
    updateOfficeHoursCheckbox();
    
    // Clear search bar when opening modal
    const searchInput = document.getElementById('employeeSearch');
    if (searchInput) {
        searchInput.value = '';
    }
    
    modal.style.display = 'flex';
    
    // Focus the search input after modal is displayed
    setTimeout(() => {
        if (searchInput) {
            searchInput.focus();
        }
    }, 100);
}

function closeFilterModal() {
    document.getElementById('employeeFilterModal').style.display = 'none';
}

function updateOfficeHoursCheckbox() {
    const checkbox = document.getElementById('officeHoursCheckbox');
    if (checkbox) {
        checkbox.checked = officeHoursOnly;
    }
}

function scrollToEmployee(employeeItem) {
    const employeeList = document.getElementById('employeeFilterList');
    if (employeeList && employeeItem) {
        // Calculate the position to scroll to center the selected employee
        const listRect = employeeList.getBoundingClientRect();
        const itemRect = employeeItem.getBoundingClientRect();
        const listScrollTop = employeeList.scrollTop;
        
        // Calculate the desired scroll position (center the item in the visible area)
        const itemRelativeTop = itemRect.top - listRect.top + listScrollTop;
        const listHeight = employeeList.clientHeight;
        const itemHeight = employeeItem.offsetHeight;
        const desiredScrollTop = itemRelativeTop - (listHeight / 2) + (itemHeight / 2);
        
        // Smooth scroll to the calculated position
        employeeList.scrollTo({
            top: Math.max(0, desiredScrollTop),
            behavior: 'smooth'
        });
    }
}

function populateEmployeeFilterList() {
    const employeeList = document.getElementById('employeeFilterList');
    let employees = Object.keys(employeeData);
    employees = sortEmployees(employees);
    
    const colors = ['color-1', 'color-2', 'color-3', 'color-4', 'color-5', 'color-6', 
                   'color-7', 'color-8', 'color-9', 'color-10', 'color-11', 'color-12'];
    
    employeeList.innerHTML = '';
    
    employees.forEach((employeeName) => {
        const originalIndex = Object.keys(employeeData).sort().indexOf(employeeName);
        const colorClass = colors[originalIndex % colors.length];
        const initials = getInitials(employeeName);
        
        const item = document.createElement('div');
        item.className = 'employee-checkbox-item';
        item.dataset.employeeName = employeeName;
        
        const isChecked = selectedEmployees.includes(employeeName);
        
        item.innerHTML = `
            <input type="checkbox" id="emp-${originalIndex}" ${isChecked ? 'checked' : ''}>
            <div class="employee-checkbox-info">
                <div class="employee-checkbox-initials ${colorClass}">${initials}</div>
                <span>${employeeName}</span>
            </div>
        `;
        
        // Add click event to the entire item
        item.addEventListener('click', (e) => {
            if (e.target.type !== 'checkbox') {
                const checkbox = item.querySelector('input[type="checkbox"]');
                checkbox.checked = !checkbox.checked;
            }
            
            // Clear search bar when checkbox is clicked
            const searchInput = document.getElementById('employeeSearch');
            if (searchInput) {
                searchInput.value = '';
                // Show all items again
                const allItems = document.querySelectorAll('.employee-checkbox-item');
                allItems.forEach(allItem => {
                    allItem.style.display = 'flex';
                });
                
                // Scroll to the clicked employee and focus search
                scrollToEmployee(item);
                // Re-focus the search input for next search
                setTimeout(() => {
                    searchInput.focus();
                }, 100);
            }
        });
        
        employeeList.appendChild(item);
    });
}

function filterEmployeeList() {
    const searchTerm = document.getElementById('employeeSearch').value.toLowerCase();
    const items = document.querySelectorAll('.employee-checkbox-item');
    
    items.forEach(item => {
        const employeeName = item.dataset.employeeName.toLowerCase();
        const matches = employeeName.includes(searchTerm);
        item.style.display = matches ? 'flex' : 'none';
    });
}

function selectAllEmployees() {
    const visibleCheckboxes = document.querySelectorAll('.employee-checkbox-item:not([style*="display: none"]) input[type="checkbox"]');
    visibleCheckboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
}

function deselectAllEmployees() {
    const visibleCheckboxes = document.querySelectorAll('.employee-checkbox-item:not([style*="display: none"]) input[type="checkbox"]');
    visibleCheckboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
}

function applyEmployeeFilter() {
    const checkedBoxes = document.querySelectorAll('.employee-checkbox-item input[type="checkbox"]:checked');
    selectedEmployees = Array.from(checkedBoxes).map(checkbox => {
        return checkbox.closest('.employee-checkbox-item').dataset.employeeName;
    });
    
    isFilterActive = selectedEmployees.length > 0 || officeHoursOnly;
    updateFilterButton();
    closeFilterModal();
    updateDisplay();
}

function clearEmployeeFilter() {
    selectedEmployees = [];
    isFilterActive = false;
    officeHoursOnly = false;
    updateFilterButton();
    updateDisplay();
}

function updateFilterButton() {
    const filterBtn = document.getElementById('filterBtn');
    if (filterBtn) {
        if (isFilterActive || officeHoursOnly) {
            filterBtn.innerHTML = '<i class="fas fa-times"></i> Clear Filter';
            filterBtn.className = 'filter-btn clear-mode';
        } else {
            filterBtn.innerHTML = 'Filter Options';
            filterBtn.className = 'filter-btn';
        }
    }
}

function updateFilterResults(displayedCount, totalCount) {
    const filterResultsText = document.getElementById('filterResultsText');
    if (filterResultsText) {
        if (isFilterActive || officeHoursOnly) {
            let filterDesc = '';
            if (selectedEmployees.length > 0 && officeHoursOnly) {
                filterDesc = ' (selected employees + office hours only)';
            } else if (selectedEmployees.length > 0) {
                filterDesc = ' (selected employees)';
            } else if (officeHoursOnly) {
                filterDesc = ' (office hours only)';
            }
            filterResultsText.textContent = `Showing ${displayedCount} of ${totalCount} employees${filterDesc}`;
        } else {
            filterResultsText.textContent = `Showing all ${displayedCount} employees`;
        }
    }
}

function jumpToToday() {
    setCurrentWeek();
    updateDisplay();
}

function toggleScheduleCollapse() {
    const unifiedContainer = document.getElementById('unifiedContainer');
    const expandableContent = document.getElementById('expandableContent');
    const collapseIcon = document.getElementById('collapseIcon');
    const expandCollapseText = document.getElementById('expandCollapseText');
    
    // Check if elements exist before trying to access them
    if (!expandableContent || !collapseIcon || !unifiedContainer || !expandCollapseText) {
        console.error('Required elements not found for collapse toggle');
        return;
    }
    
    if (expandableContent.classList.contains('collapsed')) {
        // Expand
        expandableContent.classList.remove('collapsed');
        unifiedContainer.classList.remove('collapsed');
        collapseIcon.className = 'fas fa-chevron-up';
        expandCollapseText.innerHTML = 'COLLAPSE <i class="fas fa-chevron-up" id="collapseIcon"></i>';
    } else {
        // Collapse
        expandableContent.classList.add('collapsed');
        unifiedContainer.classList.add('collapsed');
        collapseIcon.className = 'fas fa-chevron-down';
        expandCollapseText.innerHTML = 'EXPAND <i class="fas fa-chevron-down" id="collapseIcon"></i>';
    }
}

function showLoading(show) {
    const loading = document.getElementById('loading');
    const scheduleBody = document.getElementById('scheduleBody');
    const filterResults = document.getElementById('filterResults');
    const bottomSeparator = document.getElementById('bottomSeparator');
    
    if (loading) loading.style.display = show ? 'flex' : 'none';
    if (scheduleBody) scheduleBody.style.display = show ? 'none' : 'block';
    if (filterResults) filterResults.style.display = show ? 'none' : 'block';
    if (bottomSeparator) bottomSeparator.style.display = show ? 'none' : 'block';
}

function showError() {
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');
    const scheduleBody = document.getElementById('scheduleBody');
    const filterResults = document.getElementById('filterResults');
    const bottomSeparator = document.getElementById('bottomSeparator');
    
    if (loading) loading.style.display = 'none';
    if (error) error.style.display = 'flex';
    if (scheduleBody) scheduleBody.style.display = 'none';
    if (filterResults) filterResults.style.display = 'none';
    if (bottomSeparator) bottomSeparator.style.display = 'none';
}