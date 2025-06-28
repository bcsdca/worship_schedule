// Function to process task assignments
function processTaskAssignments(taskData, headerRow, excludedTasks) {
  const taskAssignments = [];
  taskData.forEach((row, i) => {
    row.forEach((name, j) => {
      if (name && j >= 4) { // Adjusted task columns
        const task = headerRow[j];
        if (!excludedTasks.includes(task)) {
          const date = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MM/dd/yyyy');
          taskAssignments.push({ name, task, date });
        }
      }
    });
  });
  return taskAssignments.sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));
}