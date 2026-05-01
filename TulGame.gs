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
    let userGrade = 1;
    
    // 1. Find user's current grade (Column C / Index 2)
    const cleanUser = username ? username.toString().trim().toLowerCase() : "";
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] && userData[i][0].toString().trim().toLowerCase() === cleanUser) { 
        userGrade = parseInt(userData[i][2]) || 1; 
        break; 
      }
    }

    const tulData = tulSheet.getDataRange().getValues();
    
    // 2. Filter and Map Deck (Single source of truth for stats and images)
    const deck = tulData.slice(1)
      .filter(row => {
        const tulName = row[0];
        const tulRequiredGrade = parseInt(row[5]);
        // Only include if pattern exists and is within user's grade
        return tulName && !isNaN(tulRequiredGrade) && tulRequiredGrade <= userGrade;
      }) 
      .map(row => ({
        name: row[0],         
        movements: parseInt(row[1]) || 0, 
        stances: parseInt(row[2]) || 0,   
        readyStance: row[3] || "---",  
        difficulty: parseInt(row[4]) || 0, 
        meaning: row[6] || "",      
        // DRY: Centralized Placeholder image
        img: row[7] || "https://placehold.co/300x200?text=No+Pattern+Image"
      }));

    if (deck.length < 4) {
      return { error: "Insufficient Tuls unlocked for your grade (" + userGrade + ")." };
    }

    // 3. Shuffle and split deck
    const shuffled = deck.sort(() => Math.random() - 0.5);
    const mid = Math.ceil(shuffled.length / 2);
    
    return {
      playerHand: shuffled.slice(0, mid),
      cpuHand: shuffled.slice(mid)
    };
  } catch (e) {
    console.error("getTulTrumpsData Error: " + e.message);
    return { error: "System Error: " + e.message };
  }
}