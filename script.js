let scheduleData = {};
let currentWeekOffset = 0;
const WEEK_1_START = new Date('2025-09-13T00:00:00');

// Generate random colors for employee initials
const colors = [
    '#e91e63', '#9c27b0', '#673ab7', '#3f51b5', '#2196f3',
    '#03a9f4', '#00bcd4', '#009688', '#4caf50', '#8bc34a',
    '#cddc39', '#ffeb3b', '#ffc107', '#ff9800', '#ff5722'
];

function getCurrentWeek() {
    const now = new Date();
    const currentDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const startDate = new Date(WEEK_1_START);
    
    // Add week offset (7 days per week, not 14)
    startDate.setDate(startDate.getDate() + (currentWeekOffset * 7));
    
    const diffTime = currentDate - startDate;
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    const weekNumber = Math.floor(diffDays / 7);
    
    return (weekNumber % 2) + 1;
}

function updateDateRange() {
    const startDate = new Date(WEEK_1_START);
    // Each week is 7 days, not 14
    startDate.setDate(startDate.getDate() + (currentWeekOffset * 7));
    
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 6); // 7 days total (0-6)
    
    const options = { month: 'long', day: 'numeric', year: 'numeric' };
    const startStr = startDate.toLocaleDateString('en-US', options);
    const endStr = endDate.toLocaleDateString('en-US', options);
    
    document.getElementById('dateRange').textContent = `${startStr} - ${endStr}`;
    
    // Update week selector
    const currentWeek = getCurrentWeek();
    document.getElementById('weekSelector').value = currentWeek;
}

function navigateWeek(direction) {
    currentWeekOffset += direction;
    updateDateRange();
    if (Object.keys(scheduleData).length > 0) {
        renderSchedule();
    }
}

function switchWeek() {
    const selectedWeek = parseInt(document.getElementById('weekSelector').value);
    const currentWeek = getCurrentWeek();
    
    if (selectedWeek !== currentWeek) {
        // Switch to the other week
        const weekDiff = selectedWeek - currentWeek;
        navigateWeek(weekDiff > 0 ? 1 : -1);
    }
}

function getColorForEmployee(name) {
    let hash = 0;
    for (let i = 0; i < name.length; i++) {
        hash = name.charCodeAt(i) + ((hash << 5) - hash);
    }
    return colors[Math.abs(hash) % colors.length];
}

function getInitials(name) {
    return name.split(' ').map(n => n[0]).join('').toUpperCase().substring(0, 2);
}

function parseTime(timeValue) {
    // Handle null, undefined, or empty values
    if (timeValue === null || timeValue === undefined || timeValue === '') {
        return null;
    }
    
    try {
        // Handle Excel decimal time format (e.g., 0.291666666666667 = 7:00 AM)
        if (typeof timeValue === 'number') {
            // Convert decimal to 24-hour time
            const totalMinutes = Math.round(timeValue * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            
            // Convert to 12-hour format
            const ampm = hours >= 12 ? 'PM' : 'AM';
            const displayHours = hours % 12 || 12;
            const displayMinutes = minutes.toString().padStart(2, '0');
            
            return `${displayHours}:${displayMinutes} ${ampm}`;
        }
        
        // Handle Date objects (if Excel returns them as dates)
        if (timeValue instanceof Date) {
            const hours = timeValue.getHours();
            const minutes = timeValue.getMinutes();
            const ampm = hours >= 12 ? 'PM' : 'AM';
            const displayHours = hours % 12 || 12;
            const displayMinutes = minutes.toString().padStart(2, '0');
            
            return `${displayHours}:${displayMinutes} ${ampm}`;
        }
        
        // Handle string values
        if (typeof timeValue === 'string') {
            const cleanTime = timeValue.trim();
            
            // If it's already in AM/PM format, return it
            if (cleanTime.includes('AM') || cleanTime.includes('PM')) {
                return cleanTime;
            }
            
            // Try to parse as 24-hour format
            const timeRegex = /^(\d{1,2}):(\d{2})$/;
            const match = cleanTime.match(timeRegex);
            
            if (match) {
                let hour = parseInt(match[1]);
                const minute = match[2];
                const ampm = hour >= 12 ? 'PM' : 'AM';
                hour = hour % 12 || 12;
                return `${hour}:${minute} ${ampm}`;
            }
        }
        
        // Try to convert to string and parse
        const stringValue = String(timeValue).trim();
        if (stringValue && stringValue !== 'null' && stringValue !== 'undefined') {
            return stringValue;
        }
        
    } catch (error) {
        console.error('Error parsing time value:', timeValue, error);
    }
    
    return null;
}

function isValidData(start, end, location) {
    return start !== null && start !== undefined && start !== '' &&
           end !== null && end !== undefined && end !== '' &&
           location !== null && location !== undefined && location !== '';
}

function loadScheduleFile() {
    // Load the TMS-WorkSchedules.xlsx file automatically
    fetch('TMS-WorkSchedules.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Could not load TMS-WorkSchedules.xlsx file');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            try {
                const workbook = XLSX.read(data, {type: 'array'});
                
                scheduleData = {};
                
                // Process each sheet
                workbook.SheetNames.forEach(sheetName => {
                    if (sheetName === 'NewEmployee' || sheetName === 'FormTools') {
                        return; // Skip these sheets
                    }
                    
                    const worksheet = workbook.Sheets[sheetName];
                    const employeeName = worksheet['C1'] ? worksheet['C1'].v : sheetName;
                    
                    const employee = {
                        name: employeeName,
                        weeks: {
                            1: { Mon: [], Tue: [], Wed: [], Thu: [], Fri: [] },
                            2: { Mon: [], Tue: [], Wed: [], Thu: [], Fri: [] }
                        }
                    };
                    
                    // Parse Week 1 (rows 9-13)
                    for (let blockNum = 1; blockNum <= 5; blockNum++) {
                        const rowNum = 8 + blockNum; // Row 9-13 in Excel = index 8-12
                        
                        const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
                        for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
                            const day = days[dayIndex];
                            const startCol = 4 + (dayIndex * 3); // Start, End, Location columns
                            const endCol = startCol + 1;
                            const locCol = startCol + 2;
                            
                            const startCell = XLSX.utils.encode_cell({r: rowNum, c: startCol});
                            const endCell = XLSX.utils.encode_cell({r: rowNum, c: endCol});
                            const locCell = XLSX.utils.encode_cell({r: rowNum, c: locCol});
                            
                            const startTime = worksheet[startCell] ? worksheet[startCell].v : null;
                            const endTime = worksheet[endCell] ? worksheet[endCell].v : null;
                            const location = worksheet[locCell] ? worksheet[locCell].v : null;
                            
                            if (isValidData(startTime, endTime, location)) {
                                const parsedStart = parseTime(startTime);
                                const parsedEnd = parseTime(endTime);
                                let parsedLocation = safeString(location).toLowerCase().trim();
                                
                                // Handle different location formats
                                if (parsedLocation.includes('remote')) parsedLocation = 'remote';
                                if (parsedLocation.includes('office')) parsedLocation = 'office';
                                
                                console.log(`Week 1 ${day}:`, {startTime, endTime, location, parsedStart, parsedEnd, parsedLocation});
                                
                                if (parsedStart && parsedEnd && (parsedLocation === 'remote' || parsedLocation === 'office')) {
                                    employee.weeks[1][day].push({
                                        start: parsedStart,
                                        end: parsedEnd,
                                        location: parsedLocation
                                    });
                                }
                            }
                        }
                    }
                    
                    // Parse Week 2 (rows 24-28)
                    for (let blockNum = 1; blockNum <= 5; blockNum++) {
                        const rowNum = 23 + blockNum; // Row 24-28 in Excel = index 23-27
                        
                        const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
                        for (let dayIndex = 0; dayIndex < days.length; dayIndex++) {
                            const day = days[dayIndex];
                            const startCol = 4 + (dayIndex * 3);
                            const endCol = startCol + 1;
                            const locCol = startCol + 2;
                            
                            const startCell = XLSX.utils.encode_cell({r: rowNum, c: startCol});
                            const endCell = XLSX.utils.encode_cell({r: rowNum, c: endCol});
                            const locCell = XLSX.utils.encode_cell({r: rowNum, c: locCol});
                            
                            const startTime = worksheet[startCell] ? worksheet[startCell].v : null;
                            const endTime = worksheet[endCell] ? worksheet[endCell].v : null;
                            const location = worksheet[locCell] ? worksheet[locCell].v : null;
                            
                            if (isValidData(startTime, endTime, location)) {
                                const parsedStart = parseTime(startTime);
                                const parsedEnd = parseTime(endTime);
                                let parsedLocation = safeString(location).toLowerCase().trim();
                                
                                // Handle different location formats
                                if (parsedLocation.includes('remote')) parsedLocation = 'remote';
                                if (parsedLocation.includes('office')) parsedLocation = 'office';
                                
                                console.log(`Week 2 ${day}:`, {startTime, endTime, location, parsedStart, parsedEnd, parsedLocation});
                                
                                if (parsedStart && parsedEnd && (parsedLocation === 'remote' || parsedLocation === 'office')) {
                                    employee.weeks[2][day].push({
                                        start: parsedStart,
                                        end: parsedEnd,
                                        location: parsedLocation
                                    });
                                }
                            }
                        }
                    }
                    
                    scheduleData[employeeName] = employee;
                });
                
                console.log('Loaded schedule data:', scheduleData);
                renderSchedule();
                
            } catch (error) {
                console.error('Parsing error:', error);
                document.getElementById('scheduleContent').innerHTML = `
                    <div class="error">
                        <div>Error parsing Excel file: ${error.message}</div>
                        <div>Please ensure the file structure matches the expected format.</div>
                    </div>
                `;
            }
        })
        .catch(error => {
            console.error('Loading error:', error);
            document.getElementById('scheduleContent').innerHTML = `
                <div class="error">
                    <div>Error loading TMS-WorkSchedules.xlsx: ${error.message}</div>
                    <div>Please ensure the file is in the same directory as this HTML file.</div>
                </div>
            `;
        });
}

function renderSchedule() {
    const currentWeek = getCurrentWeek();
    const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
    
    let html = `
        <div class="schedule-grid">
            <div class="day-header first-col"></div>
            ${dayNames.map(day => `<div class="day-header">${day}</div>`).join('')}
    `;
    
    Object.values(scheduleData).forEach(employee => {
        const color = getColorForEmployee(employee.name);
        const initials = getInitials(employee.name);
        
        html += `<div class="employee-row">`;
        html += `
            <div class="employee-name" onclick="openModal('${employee.name}')">
                <div class="employee-initial" style="background-color: ${color}">
                    ${initials}
                </div>
                <span>${employee.name}</span>
            </div>
        `;
        
        days.forEach(day => {
            const blocks = employee.weeks[currentWeek][day] || [];
            html += `<div class="day-cell">`;
            
            blocks.forEach(block => {
                const icon = block.location === 'office' ? 
                    `<svg class="work-icon" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M12 2L2 7v10c0 5.55 3.84 10 9 10s9-4.45 9-10V7L12 2z"/>
                        <path d="M12 2L2 7l10 5 10-5L12 2z"/>
                    </svg>` :
                    `<svg class="work-icon" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/>
                    </svg>`;
                
                html += `
                    <div class="work-block ${block.location}">
                        ${icon}
                        <span>${block.start} - ${block.end}</span>
                    </div>
                `;
            });
            
            html += `</div>`;
        });
        
        html += `</div>`;
    });
    
    html += `</div>`;
    
    document.getElementById('scheduleContent').innerHTML = html;
}

function openModal(employeeName) {
    const employee = scheduleData[employeeName];
    const currentWeek = getCurrentWeek();
    const otherWeek = currentWeek === 1 ? 2 : 1;
    
    document.getElementById('modalTitle').textContent = `${employee.name} - Schedule`;
    
    let modalContent = '';
    
    [currentWeek, otherWeek].forEach((weekNum, index) => {
        const weekTitle = index === 0 ? `Week ${weekNum} (Current)` : `Week ${weekNum}`;
        
        modalContent += `
            <div class="week-section">
                <h3 class="week-title">${weekTitle}</h3>
                <div class="schedule-detail">
        `;
        
        const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
        const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        
        days.forEach((day, dayIndex) => {
            const blocks = employee.weeks[weekNum][day] || [];
            
            modalContent += `
                <div class="day-detail">
                    <div class="day-detail-header">${dayNames[dayIndex]}</div>
                    <div class="day-detail-content">
            `;
            
            if (blocks.length === 0) {
                modalContent += '<div style="color: #999; font-style: italic;">No scheduled work</div>';
            } else {
                blocks.forEach(block => {
                    const icon = block.location === 'office' ? 
                        `<svg class="work-icon" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M12 2L2 7v10c0 5.55 3.84 10 9 10s9-4.45 9-10V7L12 2z"/>
                            <path d="M12 2L2 7l10 5 10-5L12 2z"/>
                        </svg>` :
                        `<svg class="work-icon" viewBox="0 0 24 24" fill="currentColor">
                            <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z"/>
                        </svg>`;
                    
                    modalContent += `
                        <div class="detail-work-block ${block.location}">
                            ${icon}
                            <span>${block.start} - ${block.end}</span>
                        </div>
                    `;
                });
            }
            
            modalContent += `
                    </div>
                </div>
            `;
        });
        
        modalContent += `
                </div>
            </div>
        `;
    });
    
    document.getElementById('modalBody').innerHTML = modalContent;
    document.getElementById('employeeModal').style.display = 'block';
}

function closeModal() {
    document.getElementById('employeeModal').style.display = 'none';
}

// Close modal when clicking outside
window.onclick = function(event) {
    const modal = document.getElementById('employeeModal');
    if (event.target === modal) {
        closeModal();
    }
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    updateDateRange();
    loadScheduleFile();
});