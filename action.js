// function to generate and download an Excel file
function exportToExcel(data) {
  // create a new Excel workbook
  const workbook = XLSX.utils.book_new();
  
  // convert the data to an Excel worksheet
  const worksheet = XLSX.utils.json_to_sheet(data);
  
  // add the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Program Schedule');
  
  // generate a file name based on the current date and time
  const fileName = `program_schedule_${new Date().toISOString()}.xlsx`;
  
  // write the workbook to a binary string
  const fileContent = XLSX.write(workbook, { type: 'binary', bookType: 'xlsx' });
  
  // create a blob from the binary string
  const blob = new Blob([s2ab(fileContent)], { type: 'application/octet-stream' });
  
  // create a download link and click it
  const downloadLink = document.createElement('a');
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = fileName;
  downloadLink.click();
}

// function to convert a string to an ArrayBuffer
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}

// function to generate the program schedule
function generateSchedule() {
  const startDate = document.getElementById('start-date').value;
  const totalWeeks = parseInt(document.getElementById('total-weeks').value);
  const moduleTitles = document.getElementById('module-titles').value.split('\n');
  
  // create an empty array to store the weeks
  const weeks = [];
  
  // create a new Date object for the start date
  const currentDate = new Date(startDate);
  
  // generate the schedule for each week
  for (let i = 1; i <= totalWeeks; i++) {
    const weekStartDate = new Date(currentDate);
    const weekEndDate = new Date(currentDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    
    const week = {
      weekNumber: i,
      startDate: weekStartDate.toISOString().substring(0, 10),
      endDate: weekEndDate.toISOString().substring(0, 10),
      moduleTitle: moduleTitles[i - 1] || '',
      nationalHolidays: []
    };
    
    weeks.push(week);
    currentDate.setDate(currentDate.getDate() + 7);
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
