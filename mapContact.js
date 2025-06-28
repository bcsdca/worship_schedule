function mapContact() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("contact");
    if (!sheet) throw new Error("Sheet 'contact' not found.");

    const data = sheet.getDataRange().getDisplayValues();

    const emailMap = new Map();
    const phoneMap = new Map();
    const carrierMap = new Map();

    data.slice(1).forEach(row => {
      const [name, email, phone, carrier] = row;
      
      // Skip empty names
      if (!name) return;
      emailMap.set(name, email || '');
      phoneMap.set(name, phone || '');
      carrierMap.set(name, carrier || '');
    });
    
    //emailMap.forEach((value, key) => {
    //  logMessage(getCallStackTrace() + `: ${key}: ${value}`);
    //});
    
    return { emailMap, phoneMap, carrierMap };

  } catch (error) {
    logMessageError(getCallStackTrace() + ": Error in mapContact: " + error.message);
    return null;
  }
}