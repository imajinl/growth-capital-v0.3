function onChangeMasterTrigger(e) {
  runAllFunctions();
}

function runAllFunctions() {
  const fns = [
    updateScoringMatchIndexInvestStats,
    updateScoringMatchIndexTags,
    updateScoringMatchIndexEcosystem,
    updateScoringMatchIndexArea,
    updateScoringMatchIndexValuation,
    updateScoringMatchIndexRaise,
    updateScoringMatchIndexRound,
    updateScoringMatchIndexTotalFunding,
    scoreInvestorsUniversalToNewTab,
    scoreInvestorsNUV,
    writeTopInvestorScores,
    writeDynamicNarrativeToC22
  ];

  fns.forEach(fn => {
    try {
      fn();
    } catch (err) {
      Logger.log(`⚠️ Error in ${fn.name}: ${err.message}`);
    }
  });
}