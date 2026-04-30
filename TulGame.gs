/**
 * TulGame.gs - Handles only the Tul Trumps logic
 */

function getTulTrumpsData(username) {
  // Use the testing default if needed
  if (!username || username === "Guest") {
    username = "Jon"; 
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tulSheet = ss.getSheetByName("Tuls");
  const userSheet = ss.getSheetByName("Users");
  
  const userData = userSheet.getDataRange().getValues();
  let userGrade = 1;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim() === username.toString().trim()) { 
      userGrade = parseInt(userData[i][2]) || 1; 
      break; 
    }
  }

  const tulData = tulSheet.getDataRange().getValues();
  const deck = tulData.slice(1)
    .filter(row => row[0] && parseInt(row[5]) <= userGrade) 
    .map(row => ({
      name: row[0],         
      movements: parseInt(row[1]) || 0, 
      stances: parseInt(row[2]) || 0,   
      readyStance: row[3],  
      difficulty: parseInt(row[4]) || 0, 
      meaning: row[6],      
      img: row[7] || "https://placehold.co/300x200?text=Pattern+Image"
    }));

  if (deck.length < 2) return { error: "Not enough Tuls unlocked for your grade!" };

  const shuffled = deck.sort(() => Math.random() - 0.5);
  const mid = Math.ceil(shuffled.length / 2);
  
  return {
    playerHand: shuffled.slice(0, mid),
    cpuHand: shuffled.slice(mid)
  };
}