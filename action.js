function calculateDates() {
  // get input values
  const startDateInput = document.getElementById('start-date');
  const breakWeeksInput = document.getElementById('break-weeks');
  const contentWeeksInput = document.getElementById('content-weeks');
  
  const startDate = startDateInput.value;
  const breakWeeks = parseInt(breakWeeksInput.value, 10);
  const contentWeeks = parseInt(contentWeeksInput.value, 10);

 // calculate start and end dates for each week
  const totalWeeks = breakWeeks + contentWeeks;
  
  const weeks = [];
  let currentDate = new Date(startDate);
  
  for (let i = 1; i <= totalWeeks; i++) {
    const weekStartDate = new Date(currentDate);
    const weekEndDate = new Date(currentDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    weeks.push({ start: weekStartDate, end: weekEndDate });
    
    currentDate = new Date(weekEndDate.getTime() );
  }
 	
  // display result
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = '';
  
  const resultTable = document.createElement('table');
  const resultTableHeaderRow = document.createElement('tr');
  resultTableHeaderRow.innerHTML = '<th>Week</th><th>Start Date</th><th>End Date</th>';
  resultTable.appendChild(resultTableHeaderRow);
  
  for (let i = 0; i < weeks.length; i++) {
    const weekNumber = i + 0;
    const weekStartDate = weeks[i].start;
    const weekEndDate = weeks[i].end;
    
    const resultTableRow = document.createElement('tr');
    resultTableRow.innerHTML = `<td>${weekNumber}</td><td>${weekStartDate.toLocaleDateString()}</td><td>${weekEndDate.toLocaleDateString()}</td>`;
    
    if (hasNationalHoliday(weekStartDate, weekEndDate)) {
      resultTableRow.classList.add('national-holiday');
    }
    
    resultTable.appendChild(resultTableRow);
  }
  
  resultDiv.appendChild(resultTable);
  resultDiv.style.display = 'block';
}

function hasNationalHoliday(startDate, endDate) {
  // replace this with your own logic to check for national holidays
  return false;
}
  
  // display the program schedule in the table
  const tableBody = document.getElementById('schedule-table-body');
  tableBody.innerHTML = '';
  weeks.forEach((week) => {
    const row = tableBody.insertRow();
    row.insertCell().textContent = week.weekNumber;
    row.insertCell().textContent = week.moduleTitle;
    row.insertCell().textContent = week.startDate;
    row.insertCell().textContent = week.endDate;
  });
  
  // show the export button
  const exportButton = document.getElementById('export-button');
  exportButton.classList.remove('d-none');
  exportButton.addEventListener('click', () => {
    exportToExcel(weeks);
  });
}

// add event listener for the generate button
const generateButton = document.getElementById('generate-button');
generateButton.addEventListener('click', generateSchedule);
