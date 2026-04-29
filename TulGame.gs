// Inside TulGame.gs
function getTulTrumpsData(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const tulSheet = ss.getSheetByName("Tuls");
  
  const userData = userSheet.getDataRange().getValues();
  let userGrade = 1; // Default to White Belt if not found

  // Find the user's current grade (Column C/Index 2)
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] == username) {
      userGrade = parseInt(userData[i][2]);
      break;
    }
  }

  // Filter Tuls: Only show Tuls up to the user's current belt level (Column F/Index 5)
  const allTuls = tulSheet.getDataRange().getValues();
  const availableTuls = allTuls.slice(1)
    .filter(row => parseInt(row[5]) <= userGrade) 
    .map(row => ({
      name: row[0],
      movements: row[1],
      stances: row[2],
      readyStance: row[3],
      difficulty: row[4],
      meaning: row[6],
      img: row[7] // Ensure these are your GitHub URLs!
    }));

  // Shuffle and deal... (use your existing shuffle logic here)
  return shuffleAndDeal(availableTuls);
}