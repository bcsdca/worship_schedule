function getTaskNames() {
    
  // Log and return the task names for the sidebar
  //taskNames will come from globalWorkerConstraints
  logMessage(`${arguments.callee.name}: Task names extracted for user selection: ${taskNames}`);

  return taskNames;
}

