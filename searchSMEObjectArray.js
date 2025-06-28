function searchSMEObjectArray(objectArray, taskName) {
  const nameArray = [];

  objectArray.forEach(obj => {
    if (obj.smeLabel === taskName) {
      nameArray.push(obj.name);
    }
  });

  return nameArray;
}