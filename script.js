// Global variables
let employeeData = {};
let currentWeekStart = null;
let currentWeekType = 1; // 1 or 2

// Week 1 starts on September 13, 2025
const WEEK_1_START = new Date('2025-09-13T00:00:00');

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    loadScheduleData();
});

function initializeEventListeners() {
    // Navigation buttons
    document.getElementById('prevWeek').addEventListener('click', () => navigateWeek(-1));
    document.getElementById('nextWeek').addEventListener('click', () => navigateWeek(1));
    
    // Week selector
    document.getElementById('weekSelector').addEventListener('change', (e) => {
        currentWeekType = parseInt(e.target.value);
        updateDisplay();
    });
    
    // Modal events
    document.getElementById('closeModal').addEventListener('click', closeModal);
    document.getElementById('employeeModal').addEventListener('click', (e) => {
        if (e.target.id === 'employeeModal') closeModal();
    });
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
        name !== 'NewEmployee' && name !== 'FormTools'
    );
    
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
    // The issue was that we were incorrectly reordering valid times
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
    
    // Update week selector
    document.getElementById('weekSelector').value = currentWeekType.toString();
}

function navigateWeek(direction) {
    currentWeekStart.setDate(currentWeekStart.getDate() + (direction * 7));
    
    // Determine week type based on weeks from original start
    const daysSinceStart = Math.floor((currentWeekStart - WEEK_1_START) / (1000 * 60 * 60 * 24));
    const weekNumber = Math.floor(daysSinceStart / 7);
    currentWeekType = (weekNumber % 2) + 1;
    
    document.getElementById('weekSelector').value = currentWeekType.toString();
    updateDisplay();
}

function updateDisplay() {
    updateDateRange();
    renderScheduleGrid();
}

function updateDateRange() {
    const weekEnd = new Date(currentWeekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    const startStr = currentWeekStart.toLocaleDateString('en-US', options);
    const endStr = weekEnd.toLocaleDateString('en-US', options);
    
    document.getElementById('dateRange').textContent = `${startStr} - ${endStr}`;
}

function renderScheduleGrid() {
    const employeeRows = document.getElementById('employeeRows');
    employeeRows.innerHTML = '';
    
    const employees = Object.keys(employeeData).sort();
    const colors = ['color-1', 'color-2', 'color-3', 'color-4', 'color-5', 'color-6', 
                   'color-7', 'color-8', 'color-9', 'color-10', 'color-11', 'color-12'];
    
    employees.forEach((employeeName, index) => {
        const employee = employeeData[employeeName];
        const weekData = currentWeekType === 1 ? employee.week1 : employee.week2;
        
        // Skip employees with no data for current week
        if (!hasScheduleData(weekData)) return;
        
        const row = document.createElement('div');
        row.className = 'employee-row';
        
        // Employee info cell
        const employeeCell = document.createElement('div');
        employeeCell.className = 'employee-info';
        employeeCell.onclick = () => openEmployeeModal(employeeName);
        
        const initials = getInitials(employeeName);
        const colorClass = colors[index % colors.length];
        
        employeeCell.innerHTML = `
            <div class="employee-initials ${colorClass}">${initials}</div>
            <div class="employee-name">${employeeName}</div>
        `;
        
        row.appendChild(employeeCell);
        
        // Day cells
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        days.forEach(day => {
            const dayCell = document.createElement('div');
            dayCell.className = 'day-cell';
            
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
            
            row.appendChild(dayCell);
        });
        
        employeeRows.appendChild(row);
    });
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

function openEmployeeModal(employeeName) {
    const modal = document.getElementById('employeeModal');
    const employee = employeeData[employeeName];
    
    document.getElementById('modalEmployeeName').textContent = `${employeeName} - Schedule`;
    
    // Update week titles
    document.getElementById('currentWeekTitle').textContent = `Week ${currentWeekType} (Current)`;
    document.getElementById('otherWeekTitle').textContent = `Week ${currentWeekType === 1 ? 2 : 1}`;
    
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

function showLoading(show) {
    document.getElementById('loading').style.display = show ? 'flex' : 'none';
    document.getElementById('scheduleBody').style.display = show ? 'none' : 'block';
}

function showError() {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('error').style.display = 'flex';
    document.getElementById('scheduleBody').style.display = 'none';
}