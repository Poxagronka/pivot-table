
function debugSingleProject() {
  const project = showChoice('Select Project:', MENU_PROJECTS);
  if (!project) return;
  
  try {
    setCurrentProject(MENU_PROJECTS[project-1].toUpperCase());
    const dateRange = getDateRange(7);
    const raw = fetchCampaignData(dateRange);
    const count = raw.data?.analytics?.richStats?.stats?.length || 0;
    
    SpreadsheetApp.getUi().alert(
      `✅ ${CURRENT_PROJECT} API works: ${count} records`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${e.toString()}`);
  }
}