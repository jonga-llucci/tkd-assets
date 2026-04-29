/** 
 * Logic for the Matching Cards game 
 */
function getMatchingData(belt, user) {
  const data = getRawSheetData(); // Shared helper in Code.gs
  const filtered = data.filter(r => r[17] === belt && r[16] !== "N" && r[18] && r[19]);
  
  // Logic to pick 6 random pairs
  const shuffled = filtered.sort(() => 0.5 - Math.random()).slice(0, 6);
  let cards = [];
  shuffled.forEach(r => {
    cards.push({ id: r[0], text: r[18], type: 'term' });
    cards.push({ id: r[0], text: r[19], type: 'def' });
  });
  
  return cards.sort(() => 0.5 - Math.random());
}