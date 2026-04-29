/** 
 * Logic for the True/False game 
 */
function getTFData(belt) {
  const data = getRawSheetData();
  const filtered = data.filter(r => r[17] === belt && r[16] !== "N");
  
  return filtered.map(r => {
    const isTrue = Math.random() > 0.5;
    let displayDef = r[19];
    
    if (!isTrue) {
      // Pick a random wrong definition from the same belt
      const others = filtered.filter(o => o[0] !== r[0]);
      displayDef = others[Math.floor(Math.random() * others.length)][19];
    }
    
    return {
      question: r[18],
      definition: displayDef,
      correct: isTrue,
      id: r[0]
    };
  }).sort(() => 0.5 - Math.random());
}