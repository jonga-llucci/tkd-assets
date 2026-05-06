/**
 * TulGame.gs
 */
function getTulTrumpsData(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userRow = ss.getSheetByName("Users").getDataRange().getValues().find(r => r[0]?.toString().toLowerCase() === username.toLowerCase());
    const userGrade = userRow ? Number(userRow[2]) : 1;

    const deck = ss.getSheetByName("Tuls").getDataRange().getValues().slice(1)
      .filter(row => row[0] && !isNaN(row[4]) && Number(row[5]) <= userGrade) 
      .map(row => ({
        name: row[0], movements: parseInt(row[1]) || 0, stances: parseInt(row[2]) || 0,
        readyStance: row[3] || "---", difficulty: parseInt(row[4]) || 0,
        meaning: row[6] || "", img: row[7] || "https://placehold.co/300x200?text=Pattern"
      }));

    if (deck.length < 4) return { error: `Unlock more Tuls! (Current Grade: ${userGrade})` };
    const shuffled = deck.sort(() => Math.random() - 0.5);
    return { playerHand: shuffled.slice(0, Math.ceil(shuffled.length/2)), cpuHand: shuffled.slice(Math.ceil(shuffled.length/2)) };
  } catch (e) { return { error: e.message }; }
}