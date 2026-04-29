function getTulTrumpsData(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userGrade = getUserGrade(username); // Shared in Code.gs
  const tulSheet = ss.getSheetByName("Tuls");
  const allTuls = tulSheet.getDataRange().getValues().slice(1);

  const availableTuls = allTuls.filter(row => parseInt(row[5]) <= userGrade) 
    .map(row => ({
      name: row[0],
      movements: row[1],
      stances: row[2],
      difficulty: row[4]
    }));

  const shuffled = availableTuls.sort(() => 0.5 - Math.random());
  return {
    playerHand: shuffled.slice(0, 5),
    cpuHand: shuffled.slice(5, 10)
  };
}