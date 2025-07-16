function createEnhancedPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);

  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  const result = buildTableData(appData, wow);
  
  writeTableData(sheet, headers, result.tableData);
  applyUnifiedFormatting(sheet, result.formatData, result.tableData.length);
  createUnifiedGrouping(sheet, result.groupData);
  
  sheet.setFrozenRows(1);
  sheet.hideColumns(1);
  
  const cache = new CommentCache();
  cache.applyCommentsToSheet();
}

function buildTableData(appData, wow) {
  const tableData = [getUnifiedHeaders()];
  const formatData = [];
  const groupData = [];
  
  const sortedApps = Object.keys(appData).sort((a, b) => 
    appData[a].appName.localeCompare(appData[b].appName));
  
  sortedApps.forEach(appKey => {
    const app = appData[appKey];
    const appRow = tableData.length;
    
    formatData.push({row: appRow + 1, type: 'APP'});
    tableData.push(['APP', app.appName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    const sortedWeeks = Object.keys(app.weeks).sort();
    let appContentRows = 0;
    
    sortedWeeks.forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekRow = tableData.length;
      formatData.push({row: weekRow + 1, type: 'WEEK'});
      
      let weekData, weekContentRows = 0;
      
      if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        weekData = processOverallWeek(week, wow, weekKey, app.appName);
        weekContentRows = Object.keys(week.networks).length;
      } else if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        weekData = processTrickyWeek(week, wow, weekKey, app.appName);
        weekContentRows = calculateTrickyRows(week.sourceApps);
      } else {
        weekData = processStandardWeek(week, wow, weekKey, app.appName);
        weekContentRows = week.campaigns ? week.campaigns.length : 0;
      }
      
      tableData.push(weekData.weekRow);
      weekData.contentRows.forEach(row => {
        formatData.push({row: tableData.length + 1, type: row.type});
        tableData.push(row.data);
      });
      
      if (weekData.groups) {
        weekData.groups.forEach(g => groupData.push({
          ...g, startRow: weekRow + 1 + g.offset
        }));
      }
      
      if (weekContentRows > 0) {
        groupData.push({
          type: 'week', startRow: weekRow + 2, count: weekContentRows
        });
      }
      
      appContentRows += 1 + weekContentRows;
    });
    
    if (appContentRows > 0) {
      groupData.push({
        type: 'app', startRow: appRow + 2, count: appContentRows
      });
    }
  });
  
  return {tableData, formatData, groupData};
}

function processOverallWeek(week, wow, weekKey, appName) {
  const networks = Object.values(week.networks || {});
  const totals = calculateNetworkTotals(networks);
  const weekWoW = wow.appWeekWoW[`${appName}_${weekKey}`] || {};
  
  const weekRow = createWeekRow(week, totals, 
    weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '',
    weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '',
    weekWoW.growthStatus || '');
  
  const contentRows = Object.keys(week.networks)
    .sort((a, b) => (week.networks[b].spend || 0) - (week.networks[a].spend || 0))
    .map(networkId => {
      const network = week.networks[networkId];
      const networkWoW = wow.networkWoW[`${networkId}_${weekKey}`] || {};
      return {
        type: 'NETWORK',
        data: createNetworkRow(network,
          networkWoW.spendChangePercent ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '',
          networkWoW.eProfitChangePercent ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '',
          networkWoW.growthStatus || '')
      };
    });
  
  return {weekRow, contentRows, groups: null};
}

function processTrickyWeek(week, wow, weekKey, appName) {
  const allCampaigns = [];
  Object.values(week.sourceApps || {}).forEach(sa => allCampaigns.push(...sa.campaigns));
  
  const totals = calculateWeekTotals(allCampaigns);
  const weekWoW = wow.appWeekWoW[`${appName}_${weekKey}`] || {};
  
  const weekRow = createWeekRow(week, totals,
    weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '',
    weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '',
    weekWoW.growthStatus || '');
  
  const contentRows = [];
  const groups = [];
  let offset = 1;
  
  const sortedSourceApps = Object.keys(week.sourceApps).sort((a, b) => {
    const spendA = week.sourceApps[a].campaigns.reduce((s, c) => s + c.spend, 0);
    const spendB = week.sourceApps[b].campaigns.reduce((s, c) => s + c.spend, 0);
    return spendB - spendA;
  });
  
  sortedSourceApps.forEach(sourceAppKey => {
    const sourceApp = week.sourceApps[sourceAppKey];
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    const sourceAppWoW = wow.sourceAppWoW[`${sourceApp.sourceAppId}_${weekKey}`] || {};
    
    contentRows.push({
      type: 'SOURCE_APP',
      data: createSourceAppRow(sourceApp.sourceAppName, sourceAppTotals,
        sourceAppWoW.spendChangePercent ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '',
        sourceAppWoW.eProfitChangePercent ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '',
        sourceAppWoW.growthStatus || '')
    });
    
    sourceApp.campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
      const campaignWoW = wow.campaignWoW[`${campaign.campaignId}_${weekKey}`] || {};
      const campaignId = (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') 
        ? `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`
        : campaign.campaignId;
      
      contentRows.push({
        type: 'CAMPAIGN',
        data: createCampaignRow(campaign, campaignId,
          campaignWoW.spendChangePercent ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '',
          campaignWoW.eProfitChangePercent ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '',
          campaignWoW.growthStatus || '')
      });
    });
    
    if (sourceApp.campaigns.length > 0) {
      groups.push({
        type: 'campaign', offset: offset + 1, count: sourceApp.campaigns.length
      });
    }
    
    offset += 1 + sourceApp.campaigns.length;
  });
  
  return {weekRow, contentRows, groups};
}

function processStandardWeek(week, wow, weekKey, appName) {
  const totals = calculateWeekTotals(week.campaigns || []);
  const weekWoW = wow.appWeekWoW[`${appName}_${weekKey}`] || {};
  
  const weekRow = createWeekRow(week, totals,
    weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '',
    weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '',
    weekWoW.growthStatus || '');
  
  const contentRows = (week.campaigns || [])
    .sort((a, b) => b.spend - a.spend)
    .map(campaign => {
      const campaignWoW = wow.campaignWoW[`${campaign.campaignId}_${weekKey}`] || {};
      const campaignId = (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR')
        ? `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`
        : campaign.campaignId;
      
      return {
        type: 'CAMPAIGN',
        data: createCampaignRow(campaign, campaignId,
          campaignWoW.spendChangePercent ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '',
          campaignWoW.eProfitChangePercent ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '',
          campaignWoW.growthStatus || '')
      };
    });
  
  return {weekRow, contentRows, groups: null};
}

function calculateTrickyRows(sourceApps) {
  return Object.values(sourceApps).reduce((total, sa) => 
    total + 1 + sa.campaigns.length, 0);
}

function writeTableData(sheet, headers, tableData) {
  sheet.getRange(1, 1, tableData.length, headers.length).setValues(tableData);
  sheet.setRowHeight(1, 40);
  TABLE_CONFIG.COLUMN_WIDTHS.forEach(col => sheet.setColumnWidth(col.c, col.w));
}

function applyUnifiedFormatting(sheet, formatData, numRows) {
  const headerRange = sheet.getRange(1, 1, 1, 19);
  headerRange.setBackground(COLORS.HEADER.background)
    .setFontColor(COLORS.HEADER.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);
  
  if (numRows > 1) {
    sheet.getRange(2, 1, numRows - 1, 19).setVerticalAlignment('middle');
    sheet.getRange(2, 19, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
    sheet.getRange(2, 18, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
    
    const rowColors = {APP: COLORS.APP_ROW, WEEK: COLORS.WEEK_ROW, 
      NETWORK: COLORS.NETWORK_ROW, SOURCE_APP: COLORS.SOURCE_APP_ROW, 
      CAMPAIGN: COLORS.CAMPAIGN_ROW};
    
    formatData.forEach(item => {
      const color = rowColors[item.type];
      if (color && item.row <= numRows) {
        const range = sheet.getRange(item.row, 1, 1, 19);
        range.setBackground(color.background);
        if (color.fontColor) range.setFontColor(color.fontColor);
        if (item.type === 'APP') range.setFontWeight('bold').setFontSize(10);
        else if (item.type !== 'WEEK') range.setFontSize(9);
      }
    });
    
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00');
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00');
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0.00');
  }
  
  applyConditionalFormatting(sheet, numRows);
}

function createUnifiedGrouping(sheet, groupData) {
  groupData.forEach(group => {
    try {
      sheet.getRange(group.startRow, 1, group.count, 1).shiftRowGroupDepth(1);
      sheet.getRange(group.startRow, 1, group.count, 1).collapseGroups();
    } catch (e) {}
  });
}

function applyConditionalFormatting(sheet, numRows) {
  const rules = [];
  
  if (numRows > 1) {
    const spendRange = sheet.getRange(2, 6, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberGreaterThan(0)
        .setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([spendRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberLessThan(0)
        .setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([spendRange]).build()
    );
    
    const profitRange = sheet.getRange(2, 17, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberGreaterThan(0)
        .setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([profitRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberLessThan(0)
        .setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([profitRange]).build()
    );
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let appName = '';
      if (data[i][0] === 'APP') {
        appName = data[i][1];
      } else {
        for (let j = i - 1; j >= 1; j--) {
          if (data[j][0] === 'APP') {
            appName = data[j][1];
            break;
          }
        }
      }
      
      const targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
      const cellRange = sheet.getRange(i + 1, 15, 1, 1);
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O${i + 1})), VALUE(SUBSTITUTE(O${i + 1},"%","")) >= ${targetEROAS})`)
          .setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor)
          .setRanges([cellRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O${i + 1})), VALUE(SUBSTITUTE(O${i + 1},"%","")) >= 120, VALUE(SUBSTITUTE(O${i + 1},"%","")) < ${targetEROAS})`)
          .setBackground(COLORS.WARNING.background).setFontColor(COLORS.WARNING.fontColor)
          .setRanges([cellRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O${i + 1})), VALUE(SUBSTITUTE(O${i + 1},"%","")) < 120)`)
          .setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor)
          .setRanges([cellRange]).build()
      );
    }
  }
  
  sheet.setConditionalFormatRules(rules);
}

function getUnifiedHeaders() {
  return ['Level','Week Range / Source App','ID','GEO','Spend','Spend WoW %','Installs','CPI','ROAS D-1','IPM','RR D-1','RR D-7','eARPU 365d','eROAS 365d','eROAS 730d','eProfit 730d','eProfit 730d WoW %','Growth Status','Comments'];
}

function createWeekRow(week, totals, spendWoW, profitWoW, status) {
  return ['WEEK', `${week.weekStart} - ${week.weekEnd}`, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    totals.avgRoas.toFixed(2), totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(1)}%`, `${totals.avgRrD7.toFixed(1)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, `${totals.avgEROASD730.toFixed(0)}%`,
    totals.totalProfit.toFixed(2), profitWoW, status, ''];
}

function createNetworkRow(network, spendWoW, profitWoW, status) {
  const installs = network.installs || 0;
  const spend = network.spend || 0;
  const cpi = installs > 0 ? spend / installs : 0;
  
  return ['NETWORK', network.networkName, '', 'ALL',
    spend.toFixed(2), spendWoW, installs, cpi.toFixed(3),
    (network.roas || 0).toFixed(2), (network.ipm || 0).toFixed(1), 
    `${(network.rrD1 || 0).toFixed(1)}%`, `${(network.rrD7 || 0).toFixed(1)}%`,
    (network.eArpuForecast || 0).toFixed(3), `${(network.eRoasForecast || 0).toFixed(0)}%`, 
    `${(network.eRoasForecastD730 || 0).toFixed(0)}%`,
    (network.eProfitForecast || 0).toFixed(2), profitWoW, status, ''];
}

function createSourceAppRow(sourceAppName, totals, spendWoW, profitWoW, status) {
  return ['SOURCE_APP', sourceAppName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    totals.avgRoas.toFixed(2), totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(1)}%`, `${totals.avgRrD7.toFixed(1)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, `${totals.avgEROASD730.toFixed(0)}%`,
    totals.totalProfit.toFixed(2), profitWoW, status, ''];
}

function createCampaignRow(campaign, campaignId, spendWoW, profitWoW, status) {
  return ['CAMPAIGN', campaign.sourceApp, campaignId, campaign.geo,
    campaign.spend.toFixed(2), spendWoW, campaign.installs, (campaign.cpi || 0).toFixed(3),
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(1)}%`, `${campaign.rrD7.toFixed(1)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitWoW, status, ''];
}

function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  const avgRoas = campaigns.length ? campaigns.reduce((s, c) => s + c.roas, 0) / campaigns.length : 0;
  const avgIpm = campaigns.length ? campaigns.reduce((s, c) => s + c.ipm, 0) / campaigns.length : 0;
  const avgRrD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD1, 0) / campaigns.length : 0;
  const avgRrD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD7, 0) / campaigns.length : 0;
  const avgArpu = campaigns.length ? campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / campaigns.length : 0;
  
  const validEROAS = campaigns.filter(c => c.eRoasForecast >= 1 && c.eRoasForecast <= 1000 && c.spend > 0);
  let avgERoas = 0;
  if (validEROAS.length > 0) {
    const totalWeighted = validEROAS.reduce((sum, c) => sum + (c.eRoasForecast * c.spend), 0);
    const totalSpendValid = validEROAS.reduce((sum, c) => sum + c.spend, 0);
    avgERoas = totalSpendValid > 0 ? totalWeighted / totalSpendValid : 0;
  }
  
  const validEROASD730 = campaigns.filter(c => c.eRoasForecastD730 >= 1 && c.eRoasForecastD730 <= 1000 && c.spend > 0);
  let avgEROASD730 = 0;
  if (validEROASD730.length > 0) {
    const totalWeightedD730 = validEROASD730.reduce((sum, c) => sum + (c.eRoasForecastD730 * c.spend), 0);
    const totalSpendValidD730 = validEROASD730.reduce((sum, c) => sum + c.spend, 0);
    avgEROASD730 = totalSpendValidD730 > 0 ? totalWeightedD730 / totalSpendValidD730 : 0;
  }
  
  const totalProfit = campaigns.reduce((s, c) => s + c.eProfitForecast, 0);
  
  return {totalSpend, totalInstalls, avgCpi, avgRoas, avgIpm, avgRrD1, avgRrD7, avgArpu, avgERoas, avgEROASD730, totalProfit};
}

function calculateNetworkTotals(networks) {
  const totalSpend = networks.reduce((s, n) => s + (n.spend || 0), 0);
  const totalInstalls = networks.reduce((s, n) => s + (n.installs || 0), 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  const avgRoas = networks.length ? networks.reduce((s, n) => s + (n.roas || 0), 0) / networks.length : 0;
  const avgRrD1 = networks.length ? networks.reduce((s, n) => s + (n.rrD1 || 0), 0) / networks.length : 0;
  const avgRrD7 = networks.length ? networks.reduce((s, n) => s + (n.rrD7 || 0), 0) / networks.length : 0;
  const avgIpm = networks.length ? networks.reduce((s, n) => s + (n.ipm || 0), 0) / networks.length : 0;
  const avgArpu = networks.length ? networks.reduce((s, n) => s + (n.eArpuForecast || 0), 0) / networks.length : 0;
  
  const validEROAS = networks.filter(n => n.eRoasForecast >= 1 && n.eRoasForecast <= 1000 && n.spend > 0);
  let avgERoas = 0;
  if (validEROAS.length > 0) {
    const totalWeighted = validEROAS.reduce((sum, n) => sum + (n.eRoasForecast * n.spend), 0);
    const totalSpendValid = validEROAS.reduce((sum, n) => sum + n.spend, 0);
    avgERoas = totalSpendValid > 0 ? totalWeighted / totalSpendValid : 0;
  }
  
  const validEROASD730 = networks.filter(n => n.eRoasForecastD730 >= 1 && n.eRoasForecastD730 <= 1000 && n.spend > 0);
  let avgEROASD730 = 0;
  if (validEROASD730.length > 0) {
    const totalWeightedD730 = validEROASD730.reduce((sum, n) => sum + (n.eRoasForecastD730 * n.spend), 0);
    const totalSpendValidD730 = validEROASD730.reduce((sum, n) => sum + n.spend, 0);
    avgEROASD730 = totalSpendValidD730 > 0 ? totalWeightedD730 / totalSpendValidD730 : 0;
  }
  
  const totalProfit = networks.reduce((s, n) => s + (n.eProfitForecast || 0), 0);
  
  return {totalSpend, totalInstalls, avgCpi, avgRoas, avgIpm, avgRrD1, avgRrD7, avgArpu, avgERoas, avgEROASD730, totalProfit};
}

function createOverallPivotTable(appData) {
  createEnhancedPivotTable(appData);
}

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    createEnhancedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}