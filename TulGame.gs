function getTulTrumpsData(username) {
  if (!username) return { error: "No user session found." };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tulSheet = ss.getSheetByName("Tuls");
  const userSheet = ss.getSheetByName("Users");
  
  // 1. Get User Grade from Users sheet
  const userData = userSheet.getDataRange().getValues();
  let userGrade = 1;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim() === username.toString().trim()) { 
      userGrade = parseInt(userData[i][2]) || 1; 
      break; 
    }
  }

  // 2. Filter Tuls by BeltLevel (Col F)
  const tulData = tulSheet.getDataRange().getValues();
  const deck = tulData.slice(1)
    .filter(row => row[0] && parseInt(row[5]) <= userGrade) // Row[5] is Col F
    .map(row => ({
      name: row[0],         // Col A
      movements: parseInt(row[1]) || 0, // Col B
      stances: parseInt(row[2]) || 0,   // Col C
      readyStance: row[3],  // Col D
      difficulty: parseInt(row[4]) || 0, // Col E
      meaning: row[6],      // Col G
      img: row[7] || "https://via.placeholder.com/150" // Col H
    }));

  if (deck.length < 2) return { error: "Not enough Tuls unlocked for your grade!" };

  // 3. Shuffle and split
  const shuffled = deck.sort(() => Math.random() - 0.5);
  const mid = Math.ceil(shuffled.length / 2);
  
  return {
    playerHand: shuffled.slice(0, mid),
    cpuHand: shuffled.slice(mid)
  };
}