function getMatchingData(username) {
  const userGrade = getUserGrade(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Questions");
  const data = sheet.getDataRange().getValues().slice(1);

  // Requirement: Filter rows with value in Column N (Index 13)[cite: 13]
  const filtered = data.filter(r => parseInt(r[17]) <= userGrade && r[13] !== "");
  
  const selected = filtered.sort(() => 0.5 - Math.random()).slice(0, 6);
  let cards = [];
  selected.forEach(r => {
    cards.push({ id: r[14], text: r[18], type: 'term' }); // Col S
    cards.push({ id: r[14], text: r[11], type: 'def' });  // Col L
  });
  
  return cards.sort(() => 0.5 - Math.random());
}