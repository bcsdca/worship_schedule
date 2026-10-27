/**********************************************************************
 * findUnassignedSlots.gs
 *
 * Find empty tasks that AI did not assign.
 **********************************************************************/

function findUnassignedSlots(scheduleRows, assignments) {

  const missing = [];

  scheduleRows.forEach(row => {

    row.emptyTasks.forEach(task => {

      // Skip tasks with no available workers
      if (task.options.length === 0) return;


      const found = assignments.some(a =>
        a.date === row.date &&
        a.columnName === task.columnName
      );


      if (!found) {

        missing.push(
          `${row.date}: ${task.columnName}`
        );

      }

    });

  });


  return missing;

}