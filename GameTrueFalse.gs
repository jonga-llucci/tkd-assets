function getTFData(username) {
  const userGrade = getUserGrade(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Questions");
  const data = sheet.getDataRange().getValues().slice(1);

  // Requirement: Filter rows with value in Column T (Index 19)[cite: 12]
  const filtered = data.filter(r => parseInt(r[17]) <= userGrade && r[19] !== "");
  
  return filtered.map(r => {
    const isTrue = Math.random() > 0.5;
    return {
      question: r[18], // Col S
      displayDef: isTrue ? r[11] : "DECOY_TEXT",
      isCorrect: isTrue,
      qId: r[14].toString()
    };
  }).sort(() => 0.5 - Math.random()).slice(0, 10);
}