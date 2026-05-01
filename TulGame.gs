/**
 * TulGame.gs - Handles only the Tul Trumps logic
 */

/**
 * Consolidated Tul Trumps Data Fetcher
 * Location: TulGame.gs
 */
function getTulTrumpsData(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tulSheet = ss.getSheetByName("Tuls");
    const userSheet = ss.getSheetByName("Users");
    
    const userData = userSheet.getDataRange().getValues();
    let userGrade = 1; // Default
    
    const cleanUser = username ? username.toString().trim().toLowerCase() : "";
    
    // Find user's current grade (Column C / Index 2)
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanUser) { 
        userGrade = Number(userData[i][2]); // Force numeric[cite: 3]
        break; 
      }
    }

    const tulData = tulSheet.getDataRange().getValues();
    
    const deck = tulData.slice(1)
      .filter(row => {
        const tulName = row[0];
        const tulRequiredGrade = Number(row[5]); // Column F[cite: 3]
        
        // Filter: Must have a name AND required grade <= user's grade[cite: 3]
        return tulName && !isNaN(tulRequiredGrade) && tulRequiredGrade <= userGrade;
      }) 
      .map(row => ({
        name: row[0],         
        movements: parseInt(row[1]) || 0, 
        stances: parseInt(row[2]) || 0,   
        readyStance: row[3] || "---",  
        difficulty: parseInt(row[4]) || 0, 
        meaning: row[6] || "",      
        img: row[7] || "https://placehold.co/300x200?text=No+Pattern+Image"
      }));

    if (deck.length < 4) {
      return { error: "Not enough Tuls unlocked! Grade: " + userGrade + " (Found: " + deck.length + ")" };
    }

    // Shuffle and split
    const shuffled = deck.sort(() => Math.random() - 0.5);
    const mid = Math.ceil(shuffled.length / 2);
    
    return {
      playerHand: shuffled.slice(0, mid),
      cpuHand: shuffled.slice(mid)
    };
  } catch (e) {
    return { error: e.message };
  }
}