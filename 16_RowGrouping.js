function createUnifiedRowGrouping(sheet, formatData) {
  const startTime = Date.now();
  
  if (!formatData || formatData.length === 0) {
    console.log('No format data for grouping');
    return;
  }
  
  try {
    if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
      createApplovinTestRowGrouping(sheet, formatData);
    } else if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      createIncentTrafficRowGrouping(sheet, formatData);
    } else {
      createStandardRowGrouping(sheet, formatData);
    }
    
    console.log(`Row grouping completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
  } catch (e) {
    console.error('Error creating row grouping:', e);
  }
}

function createApplovinTestRowGrouping(sheet, formatData) {
  const appGroups = new Map();
  const campaignGroups = new Map();
  
  formatData.forEach(item => {
    switch (item.type) {
      case 'APP':
        if (!appGroups.has(item.row)) {
          appGroups.set(item.row, { start: item.row, campaigns: [] });
        }
        break;
      case 'CAMPAIGN':
        const appGroup = findCurrentAppGroup(appGroups, item.row);
        if (appGroup) {
          const campaignGroup = { start: item.row, weeks: [] };
          appGroup.campaigns.push(campaignGroup);
          campaignGroups.set(item.row, campaignGroup);
        }
        break;
      case 'WEEK':
        const campaignGroup = findCurrentCampaignGroup(campaignGroups, item.row);
        if (campaignGroup) {
          campaignGroup.weeks.push(item.row);
        }
        break;
    }
  });
  
  appGroups.forEach(appGroup => {
    if (appGroup.campaigns.length > 0) {
      const lastCampaign = appGroup.campaigns[appGroup.campaigns.length - 1];
      const lastWeekRow = lastCampaign.weeks.length > 0 ? 
                         Math.max(...lastCampaign.weeks) : lastCampaign.start;
      
      if (lastWeekRow > appGroup.start) {
        try {
          const group = sheet.getRange(appGroup.start + 1, 1, lastWeekRow - appGroup.start, 1).getGroup();
          if (!group) {
            sheet.getRange(appGroup.start + 1, 1, lastWeekRow - appGroup.start, 1).createGroup();
          }
        } catch (e) {
          console.error(`Error creating app group starting at row ${appGroup.start}:`, e);
        }
      }
      
      appGroup.campaigns.forEach(campaign => {
        if (campaign.weeks.length > 0) {
          const lastWeekRow = Math.max(...campaign.weeks);
          if (lastWeekRow > campaign.start) {
            try {
              const group = sheet.getRange(campaign.start + 1, 1, lastWeekRow - campaign.start, 1).getGroup();
              if (!group) {
                sheet.getRange(campaign.start + 1, 1, lastWeekRow - campaign.start, 1).createGroup();
              }
            } catch (e) {
              console.error(`Error creating campaign group starting at row ${campaign.start}:`, e);
            }
          }
        }
      });
    }
  });
  
  try {
    collapseGroups(sheet, 1);
    console.log(`Created ${appGroups.size} app groups with campaign subgroups`);
  } catch (e) {
    console.error('Error collapsing APPLOVIN_TEST groups:', e);
  }
}

function findCurrentAppGroup(appGroups, rowNumber) {
  let currentApp = null;
  for (const [startRow, group] of appGroups) {
    if (startRow < rowNumber) {
      currentApp = group;
    } else {
      break;
    }
  }
  return currentApp;
}

function findCurrentCampaignGroup(campaignGroups, rowNumber) {
  let currentCampaign = null;
  for (const [startRow, group] of campaignGroups) {
    if (startRow < rowNumber) {
      currentCampaign = group;
    } else {
      break;
    }
  }
  return currentCampaign;
}

function createIncentTrafficRowGrouping(sheet, formatData) {
  const networkGroups = new Map();
  const weekGroups = new Map();
  
  formatData.forEach(item => {
    switch (item.type) {
      case 'NETWORK':
        networkGroups.set(item.row, { start: item.row, weeks: [] });
        break;
      case 'WEEK':
        const networkGroup = findCurrentNetworkGroup(networkGroups, item.row);
        if (networkGroup) {
          const weekGroup = { start: item.row, apps: [] };
          networkGroup.weeks.push(weekGroup);
          weekGroups.set(item.row, weekGroup);
        }
        break;
      case 'APP':
        const weekGroup = findCurrentWeekGroup(weekGroups, item.row);
        if (weekGroup) {
          weekGroup.apps.push(item.row);
        }
        break;
    }
  });
  
  networkGroups.forEach(networkGroup => {
    if (networkGroup.weeks.length > 0) {
      const lastWeek = networkGroup.weeks[networkGroup.weeks.length - 1];
      const lastAppRow = lastWeek.apps.length > 0 ? 
                        Math.max(...lastWeek.apps) : lastWeek.start;
      
      if (lastAppRow > networkGroup.start) {
        try {
          const group = sheet.getRange(networkGroup.start + 1, 1, lastAppRow - networkGroup.start, 1).getGroup();
          if (!group) {
            sheet.getRange(networkGroup.start + 1, 1, lastAppRow - networkGroup.start, 1).createGroup();
          }
        } catch (e) {
          console.error(`Error creating network group starting at row ${networkGroup.start}:`, e);
        }
      }
      
      networkGroup.weeks.forEach(week => {
        if (week.apps.length > 0) {
          const lastAppRow = Math.max(...week.apps);
          if (lastAppRow > week.start) {
            try {
              const group = sheet.getRange(week.start + 1, 1, lastAppRow - week.start, 1).getGroup();
              if (!group) {
                sheet.getRange(week.start + 1, 1, lastAppRow - week.start, 1).createGroup();
              }
            } catch (e) {
              console.error(`Error creating week group starting at row ${week.start}:`, e);
            }
          }
        }
      });
    }
  });
  
  try {
    collapseGroups(sheet, 1);
    console.log(`Created ${networkGroups.size} network groups with week subgroups`);
  } catch (e) {
    console.error('Error collapsing INCENT_TRAFFIC groups:', e);
  }
}

function findCurrentNetworkGroup(networkGroups, rowNumber) {
  let currentNetwork = null;
  for (const [startRow, group] of networkGroups) {
    if (startRow < rowNumber) {
      currentNetwork = group;
    } else {
      break;
    }
  }
  return currentNetwork;
}

function findCurrentWeekGroup(weekGroups, rowNumber) {
  let currentWeek = null;
  for (const [startRow, group] of weekGroups) {
    if (startRow < rowNumber) {
      currentWeek = group;
    } else {
      break;
    }
  }
  return currentWeek;
}

function createStandardRowGrouping(sheet, formatData) {
  const appGroups = new Map();
  const weekGroups = new Map();
  const sourceAppGroups = new Map();
  
  formatData.forEach(item => {
    switch (item.type) {
      case 'APP':
        appGroups.set(item.row, { start: item.row, weeks: [] });
        break;
      case 'WEEK':
        const appGroup = findCurrentAppGroup(appGroups, item.row);
        if (appGroup) {
          const weekGroup = { start: item.row, sourceApps: [], campaigns: [] };
          appGroup.weeks.push(weekGroup);
          weekGroups.set(item.row, weekGroup);
        }
        break;
      case 'SOURCE_APP':
        const weekGroup = findCurrentWeekGroup(weekGroups, item.row);
        if (weekGroup) {
          const sourceAppGroup = { start: item.row, campaigns: [] };
          weekGroup.sourceApps.push(sourceAppGroup);
          sourceAppGroups.set(item.row, sourceAppGroup);
        }
        break;
      case 'CAMPAIGN':
        const sourceAppGroup = findCurrentSourceAppGroup(sourceAppGroups, item.row);
        if (sourceAppGroup) {
          sourceAppGroup.campaigns.push(item.row);
        } else {
          const weekGroupForCampaign = findCurrentWeekGroup(weekGroups, item.row);
          if (weekGroupForCampaign) {
            weekGroupForCampaign.campaigns.push(item.row);
          }
        }
        break;
      case 'NETWORK':
        const weekGroupForNetwork = findCurrentWeekGroup(weekGroups, item.row);
        if (weekGroupForNetwork) {
          weekGroupForNetwork.campaigns.push(item.row);
        }
        break;
    }
  });
  
  appGroups.forEach(appGroup => {
    if (appGroup.weeks.length > 0) {
      const lastWeek = appGroup.weeks[appGroup.weeks.length - 1];
      const lastRowInWeek = getLastRowInWeek(lastWeek);
      
      if (lastRowInWeek > appGroup.start) {
        try {
          const group = sheet.getRange(appGroup.start + 1, 1, lastRowInWeek - appGroup.start, 1).getGroup();
          if (!group) {
            sheet.getRange(appGroup.start + 1, 1, lastRowInWeek - appGroup.start, 1).createGroup();
          }
        } catch (e) {
          console.error(`Error creating app group starting at row ${appGroup.start}:`, e);
        }
      }
      
      appGroup.weeks.forEach(week => {
        const lastRowInThisWeek = getLastRowInWeek(week);
        if (lastRowInThisWeek > week.start) {
          try {
            const group = sheet.getRange(week.start + 1, 1, lastRowInThisWeek - week.start, 1).getGroup();
            if (!group) {
              sheet.getRange(week.start + 1, 1, lastRowInThisWeek - week.start, 1).createGroup();
            }
          } catch (e) {
            console.error(`Error creating week group starting at row ${week.start}:`, e);
          }
        }
        
        week.sourceApps.forEach(sourceApp => {
          if (sourceApp.campaigns.length > 0) {
            const lastCampaignRow = Math.max(...sourceApp.campaigns);
            if (lastCampaignRow > sourceApp.start) {
              try {
                const group = sheet.getRange(sourceApp.start + 1, 1, lastCampaignRow - sourceApp.start, 1).getGroup();
                if (!group) {
                  sheet.getRange(sourceApp.start + 1, 1, lastCampaignRow - sourceApp.start, 1).createGroup();
                }
              } catch (e) {
                console.error(`Error creating source app group starting at row ${sourceApp.start}:`, e);
              }
            }
          }
        });
      });
    }
  });
  
  try {
    collapseGroups(sheet, 1);
    const totalGroups = appGroups.size + 
                       Array.from(appGroups.values()).reduce((sum, app) => sum + app.weeks.length, 0) +
                       Array.from(appGroups.values()).reduce((sum, app) => 
                         sum + app.weeks.reduce((weekSum, week) => weekSum + week.sourceApps.length, 0), 0);
    console.log(`Created ${totalGroups} groups (apps, weeks, source apps)`);
  } catch (e) {
    console.error('Error collapsing standard groups:', e);
  }
}

function findCurrentSourceAppGroup(sourceAppGroups, rowNumber) {
  let currentSourceApp = null;
  for (const [startRow, group] of sourceAppGroups) {
    if (startRow < rowNumber) {
      currentSourceApp = group;
    } else {
      break;
    }
  }
  return currentSourceApp;
}

function getLastRowInWeek(week) {
  let lastRow = week.start;
  
  if (week.sourceApps && week.sourceApps.length > 0) {
    week.sourceApps.forEach(sourceApp => {
      if (sourceApp.campaigns.length > 0) {
        lastRow = Math.max(lastRow, Math.max(...sourceApp.campaigns));
      } else {
        lastRow = Math.max(lastRow, sourceApp.start);
      }
    });
  }
  
  if (week.campaigns && week.campaigns.length > 0) {
    lastRow = Math.max(lastRow, Math.max(...week.campaigns));
  }
  
  return lastRow;
}

function collapseGroups(sheet, depth = 1) {
  try {
    const maxRows = sheet.getMaxRows();
    const range = sheet.getRange(1, 1, maxRows, 1);
    
    for (let level = 1; level <= depth; level++) {
      try {
        range.collapseGroups();
      } catch (e) {
        console.log(`No groups to collapse at level ${level}`);
        break;
      }
    }
  } catch (e) {
    console.error('Error in collapseGroups:', e);
  }
}

function expandGroups(sheet, depth = 1) {
  try {
    const maxRows = sheet.getMaxRows();
    const range = sheet.getRange(1, 1, maxRows, 1);
    
    for (let level = 1; level <= depth; level++) {
      try {
        range.expandGroups();
      } catch (e) {
        console.log(`No groups to expand at level ${level}`);
        break;
      }
    }
  } catch (e) {
    console.error('Error in expandGroups:', e);
  }
}

function removeAllGroups(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    const range = sheet.getRange(1, 1, maxRows, 1);
    
    let hasGroups = true;
    while (hasGroups) {
      try {
        range.expandGroups();
        const groups = range.getRowGroups();
        if (groups.length === 0) {
          hasGroups = false;
        } else {
          groups.forEach(group => {
            try {
              group.remove();
            } catch (e) {
              console.error('Error removing individual group:', e);
            }
          });
        }
      } catch (e) {
        hasGroups = false;
      }
    }
    
    console.log('All row groups removed');
  } catch (e) {
    console.error('Error removing all groups:', e);
  }
}

function createCustomGrouping(sheet, groupConfig) {
  try {
    removeAllGroups(sheet);
    
    groupConfig.forEach(config => {
      const { startRow, endRow, collapse = true } = config;
      
      if (startRow < endRow && startRow > 0) {
        try {
          const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 1);
          const group = range.createGroup();
          
          if (collapse) {
            group.collapse();
          }
        } catch (e) {
          console.error(`Error creating custom group ${startRow}-${endRow}:`, e);
        }
      }
    });
    
    console.log(`Created ${groupConfig.length} custom groups`);
  } catch (e) {
    console.error('Error in createCustomGrouping:', e);
  }
}

function optimizeGrouping(sheet, formatData) {
  const startTime = Date.now();
  
  try {
    removeAllGroups(sheet);
    createUnifiedRowGrouping(sheet, formatData);
    
    console.log(`Grouping optimization completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
  } catch (e) {
    console.error('Error in optimizeGrouping:', e);
  }
}

function toggleGroupExpansion(sheet, groupLevel = 1) {
  try {
    const maxRows = sheet.getMaxRows();
    const range = sheet.getRange(1, 1, maxRows, 1);
    
    const groups = range.getRowGroups();
    if (groups.length === 0) {
      console.log('No groups found to toggle');
      return;
    }
    
    const firstGroup = groups[0];
    const isCollapsed = firstGroup.isCollapsed();
    
    if (isCollapsed) {
      expandGroups(sheet, groupLevel);
      console.log('Groups expanded');
    } else {
      collapseGroups(sheet, groupLevel);
      console.log('Groups collapsed');
    }
  } catch (e) {
    console.error('Error toggling group expansion:', e);
  }
}

function getGroupingStats(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    const range = sheet.getRange(1, 1, maxRows, 1);
    const groups = range.getRowGroups();
    
    const stats = {
      totalGroups: groups.length,
      collapsedGroups: 0,
      expandedGroups: 0,
      groupLevels: new Set()
    };
    
    groups.forEach(group => {
      if (group.isCollapsed()) {
        stats.collapsedGroups++;
      } else {
        stats.expandedGroups++;
      }
      stats.groupLevels.add(group.getDepth());
    });
    
    stats.maxDepth = stats.groupLevels.size > 0 ? Math.max(...stats.groupLevels) : 0;
    
    console.log('Grouping stats:', stats);
    return stats;
  } catch (e) {
    console.error('Error getting grouping stats:', e);
    return { totalGroups: 0, collapsedGroups: 0, expandedGroups: 0, maxDepth: 0 };
  }
}