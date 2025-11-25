const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { getISOWeek, startOfWeek, format: formatDate } = require('date-fns');
// Fixed import for date-fns-tz - using parseISO from date-fns instead
const { parseISO } = require('date-fns');

// ====================================================================================
// CONFIGURATIONS
// ====================================================================================

const FLOW_A_CONFIG = {
    displayName: "Yieldlab Report",
    sheetName: 'Yieldlab Report', headerRowIndex: 5, groupingColumn: 'Partnership', dateColumn: 'Start',
    metrics: ['Impressions', 'Net Revenue', 'Net eCPM'],
    vtrConfig: { NUMERATOR_COLUMN: '100% played', DENOMINATOR_COLUMN: 'Impressions' }
};

const FLOW_B_CONFIG = {
    displayName: "Type B Report",
    sheetName: 'report', headerRowIndex: 0, groupingColumn: 'Deal', dateColumn: 'Day',
    metrics: ['Imps Resold', 'Video Completion Rate', 'Reseller Revenue'],
    vtrConfig: null
};

const EMOJIS = { 'Impressions': '👁️', 'Imps Resold': '🚀', 'Revenue': '💰', 'eCPM': '📈', 'VTR': '▶️', 'Rate': '📊', 'DEFAULT': '🔹' };

const ICONS = {
    'Impressions': `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path d="M12 15a3 3 0 100-6 3 3 0 000 6z"/><path fill-rule="evenodd" d="M1.323 11.447C2.811 6.976 7.028 3.75 12.001 3.75c4.97 0 9.185 3.223 10.675 7.69.12.362.12.752 0 1.113-1.487 4.471-5.701 7.697-10.672 7.697-4.97 0-9.186-3.223-10.675-7.69a.75.75 0 010-1.113zM12.001 18C8.97 18 6.13 16.333 4.172 13.5h15.65D17.87 16.333 15.032 18 12.001 18z" clip-rule="evenodd"/></svg>`,
    'Net Revenue': `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path d="M12 7.5a2.25 2.25 0 100 4.5 2.25 2.25 0 000-4.5z"/><path fill-rule="evenodd" d="M1.5 4.5a3 3 0 013-3h15a3 3 0 013 3v15a3 3 0 01-3 3h-15a3 3 0 01-3-3v-15zm4.125 3a2.25 2.25 0 100 4.5 2.25 2.25 0 000-4.5zm11.25 4.5a2.25 2.25 0 114.5 0 2.25 2.25 0 01-4.5 0zM5.625 13.5a2.25 2.25 0 100 4.5 2.25 2.25 0 000-4.5z" clip-rule="evenodd"/></svg>`,
    'Net eCPM': `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path fill-rule="evenodd" d="M12.963 2.286a.75.75 0 00-1.071 1.071L12 3.428V7.5a.75.75 0 001.5 0V4.53l1.72 1.72a.75.75 0 101.06-1.06l-3.32-3.32zM11.25 7.5a.75.75 0 01.75-.75h4.5a.75.75 0 010 1.5h-4.5a.75.75 0 01-.75-.75z" clip-rule="evenodd"/><path d="M14.25 10.5a4.5 4.5 0 11-9 0 4.5 4.5 0 019 0zM12 12.75a2.25 2.25 0 100-4.5 2.25 2.25 0 000 4.5z"/><path fill-rule="evenodd" d="M12 21a8.25 8.25 0 008.25-8.25c0-1.04-.195-2.036-.559-2.952a.75.75 0 00-1.325.755 6.75 6.75 0 01.484 2.197 6.75 6.75 0 01-13.5 0A6.75 6.75 0 019.12 7.303a.75.75 0 00-1.325-.755 8.25 8.25 0 00-2.045 5.202A8.25 8.25 0 0012 21z" clip-rule="evenodd"/></svg>`,
    'VTR': `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path fill-rule="evenodd" d="M2.25 12c0-5.385 4.365-9.75 9.75-9.75s9.75 4.365 9.75 9.75-4.365 9.75-9.75 9.75S2.25 17.385 2.25 12zm14.024-.983a1.125 1.125 0 010 1.966l-5.603 3.048a1.125 1.125 0 01-1.631-.983V8.935a1.125 1.125 0 011.631-.983l5.603 3.048z" clip-rule="evenodd"/></svg>`,
    'DEFAULT': `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path fill-rule="evenodd" d="M12 2.25c-5.385 0-9.75 4.365-9.75 9.75s4.365 9.75 9.75 9.75 9.75-4.365 9.75-9.75S17.385 2.25 12 2.25zm0 8.625a1.125 1.125 0 100 2.25 1.125 1.125 0 000-2.25zM15.375 12a3.375 3.375 0 11-6.75 0 3.375 3.375 0 016.75 0z" clip-rule="evenodd" /></svg>`
};

// ====================================================================================
// HELPER FUNCTIONS
// ====================================================================================

function getEmojiForMetric(metric) {
    for (const key in EMOJIS) { if (metric.toLowerCase().includes(key.toLowerCase())) return EMOJIS[key]; }
    return EMOJIS['DEFAULT'];
}

function getIconForMetric(metric) {
    for (const key in ICONS) {
        if (metric.toLowerCase().includes('revenue')) return ICONS['Net Revenue'];
        if (metric.toLowerCase().includes('ecpm')) return ICONS['Net eCPM'];
        if (metric.toLowerCase().includes(key.toLowerCase())) return ICONS[key];
    }
    return ICONS['DEFAULT'];
}

function getDateKey(rawDateValue) {
    if (!rawDateValue) return null;
    let dateObject;
    if (rawDateValue instanceof Date) {
        dateObject = rawDateValue;
    } else if (typeof rawDateValue === 'string') {
        const datePart = rawDateValue.substring(0, 10);
        const [year, month, day] = datePart.split('-').map(Number);
        if (year && month && day) {
            dateObject = new Date(Date.UTC(year, month - 1, day));
        }
    }
    if (!dateObject || isNaN(dateObject.getTime())) return null;
    return formatDate(dateObject, 'dd/MM/yyyy');
}

// ====================================================================================
// ANALYSIS FUNCTIONS
// ====================================================================================

function analyzeFile(filePath, config, isSingleMode) {
    const workbook = xlsx.readFile(filePath, { cellDates: true });
    const worksheet = workbook.Sheets[config.sheetName];
    if (!worksheet) throw new Error(`Sheet "${config.sheetName}" not found. Please check the Flow Type and file format.`);
    const data = xlsx.utils.sheet_to_json(worksheet, { range: config.headerRowIndex });
    if (data.length === 0) throw new Error("No data rows found in the sheet.");
    
    const metricsToProcess = config.vtrConfig ? new Set([...config.metrics, config.vtrConfig.NUMERATOR_COLUMN, config.vtrConfig.DENOMINATOR_COLUMN]) : new Set(config.metrics);
    
    if (isSingleMode) {
        const dataByDate = {};
        data.forEach(row => {
            const dateKey = getDateKey(row[config.dateColumn]);
            if (!dateKey) return;
            const groupKey = row[config.groupingColumn];
            if (!groupKey) return;
            if (!dataByDate[dateKey]) dataByDate[dateKey] = {};
            if (!dataByDate[dateKey][groupKey]) dataByDate[dateKey][groupKey] = {};
            metricsToProcess.forEach(col => {
                if (row[col] !== undefined) {
                    const v = parseFloat(row[col]);
                    if (!isNaN(v)) dataByDate[dateKey][groupKey][col] = (dataByDate[dateKey][groupKey][col] || 0) + v;
                }
            });
        });
        if (config.vtrConfig) {
            for (const dateKey in dataByDate) {
                for (const groupKey in dataByDate[dateKey]) {
                    const group = dataByDate[dateKey][groupKey];
                    const num = group[config.vtrConfig.NUMERATOR_COLUMN] || 0;
                    const den = group[config.vtrConfig.DENOMINATOR_COLUMN] || 0;
                    group['VTR'] = (den !== 0) ? (num / den) : 0;
                }
            }
        }
        return dataByDate;
    } else {
        const fileDate = getDateKey(data[0][config.dateColumn]) || `File (${path.basename(filePath)})`;
        const groupedData = {};
        data.forEach(row => {
            const groupKey = row[config.groupingColumn];
            if (!groupKey) return;
            if (!groupedData[groupKey]) groupedData[groupKey] = {};
            metricsToProcess.forEach(col => {
                if (row[col] !== undefined) {
                    const v = parseFloat(row[col]);
                    if (!isNaN(v)) groupedData[groupKey][col] = (groupedData[groupKey][col] || 0) + v;
                }
            });
        });
        if (config.vtrConfig) {
            for (const groupKey in groupedData) {
                const group = groupedData[groupKey];
                const num = group[config.vtrConfig.NUMERATOR_COLUMN] || 0;
                const den = group[config.vtrConfig.DENOMINATOR_COLUMN] || 0;
                group['VTR'] = (den !== 0) ? (num / den) : 0;
            }
        }
        return { groupedData, fileDate };
    }
}

function analyzeWeeklySingleFile(filePath, config) {
    try {
        const workbook = xlsx.readFile(filePath, { cellDates: true });
        const worksheet = workbook.Sheets[config.sheetName];
        if (!worksheet) throw new Error(`Sheet "${config.sheetName}" not found.`);
        const data = xlsx.utils.sheet_to_json(worksheet, { range: config.headerRowIndex });
        if (data.length === 0) throw new Error("No data rows found.");

        console.log(`[DEBUG] Starting weekly analysis for ${path.basename(filePath)}. Total rows found: ${data.length}`);

        const dataByWeek = {};
        const uniqueDaysPerWeek = {};
        const metricsToProcess = config.vtrConfig ? new Set([...config.metrics, config.vtrConfig.NUMERATOR_COLUMN, config.vtrConfig.DENOMINATOR_COLUMN]) : new Set(config.metrics);

        data.forEach((row) => {
            const rawDateValue = row[config.dateColumn];
            if (!rawDateValue) return;

            let dateObj;
            if (rawDateValue instanceof Date) {
                dateObj = rawDateValue;
            } else if (typeof rawDateValue === 'string') {
                const datePart = rawDateValue.substring(0, 10);
                // Use parseISO instead of zonedTimeToUtc for simpler date parsing
                dateObj = parseISO(datePart);
            } else {
                return;
            }
            
            if (!dateObj || isNaN(dateObj.getTime())) return;
            
            const weekKey = `${dateObj.getUTCFullYear()}-W${getISOWeek(dateObj).toString().padStart(2, '0')}`;
            const dateISO = dateObj.toISOString().split('T')[0];

            if (!dataByWeek[weekKey]) {
                dataByWeek[weekKey] = {};
                uniqueDaysPerWeek[weekKey] = new Set();
            }
            uniqueDaysPerWeek[weekKey].add(dateISO);
            
            const groupKey = row[config.groupingColumn];
            if (!groupKey) return;
            if (!dataByWeek[weekKey][groupKey]) dataByWeek[weekKey][groupKey] = {};

            metricsToProcess.forEach(col => {
                if (row[col] !== undefined) {
                    const v = parseFloat(row[col]);
                    if (!isNaN(v)) dataByWeek[weekKey][groupKey][col] = (dataByWeek[weekKey][groupKey][col] || 0) + v;
                }
            });
        });

        console.log('[DEBUG] Days found per week:');
        for (const week in uniqueDaysPerWeek) {
            console.log(`  - ${week}: ${uniqueDaysPerWeek[week].size} unique day(s) -> ${Array.from(uniqueDaysPerWeek[week]).sort().join(', ')}`);
        }

        const fullWeeks = Object.keys(dataByWeek).filter(weekKey => uniqueDaysPerWeek[weekKey].size === 7);
        if (fullWeeks.length < 2) {
            throw new Error(`Comparison failed. Only found ${fullWeeks.length} full week(s) of data (a full week requires 7 unique days of data).`);
        }
        
        fullWeeks.sort();
        const week2Key = fullWeeks[fullWeeks.length - 1];
        const week1Key = fullWeeks[fullWeeks.length - 2];
        
        const dateConverter = (weekStr) => {
            const [year, weekNum] = weekStr.split('-W');
            let d = new Date(Date.UTC(year, 0, 1 + (weekNum - 1) * 7));
            let day = d.getUTCDay();
            let diff = d.getUTCDate() - day + (day === 0 ? -6 : 1);
            return new Date(d.setUTCDate(diff));
        };
        const week1StartDate = dateConverter(week1Key);
        const week2StartDate = dateConverter(week2Key);
        const date1Label = `${formatDate(week1StartDate, 'dd/MM')} - ${formatDate(new Date(week1StartDate).setUTCDate(week1StartDate.getUTCDate() + 6), 'dd/MM')}`;
        const date2Label = `${formatDate(week2StartDate, 'dd/MM')} - ${formatDate(new Date(week2StartDate).setUTCDate(week2StartDate.getUTCDate() + 6), 'dd/MM')}`;

        [week1Key, week2Key].forEach(weekKey => {
            if (config.vtrConfig) {
                for (const groupKey in dataByWeek[weekKey]) {
                    const group = dataByWeek[weekKey][groupKey];
                    const num = group[config.vtrConfig.NUMERATOR_COLUMN] || 0;
                    const den = group[config.vtrConfig.DENOMINATOR_COLUMN] || 0;
                    group['VTR'] = (den !== 0) ? (num / den) : 0;
                }
            }
        });

        return {
            results1: dataByWeek[week1Key],
            results2: dataByWeek[week2Key],
            date1: date1Label,
            date2: date2Label,
            fileInfo: `Comparing weeks ${date1Label} and ${date2Label} within ${path.basename(filePath)}`
        };
    } catch (error) {
        throw new Error(`Failed during weekly analysis of "${path.basename(filePath)}": ${error.message}`);
    }
}

// ====================================================================================
// REPORT GENERATORS
// ====================================================================================

function generateConciseTextReport(results1, results2, date1, date2, config) {
    const allGroupKeys = new Set([...Object.keys(results1), ...Object.keys(results2)]);
    let output = [`--- Comparative Analysis: ${config.displayName} ---`, `Comparing ${date1} (previous) vs ${date2} (current)`];
    allGroupKeys.forEach(groupKey => {
        output.push(`\n========================================`);
        output.push(`  ${config.groupingColumn}: ${groupKey}`);
        output.push(`========================================`);
        const metricsToCompare = [...config.metrics];
        if (config.vtrConfig) metricsToCompare.push('VTR');
        metricsToCompare.forEach(metric => {
            const val1 = (results1[groupKey] || {})[metric] || 0;
            const val2 = (results2[groupKey] || {})[metric] || 0;
            const difference = val2 - val1;
            const percentageChange = (val1 !== 0) ? (difference / val1) * 100 : Infinity;
            const isVTR = metric === 'VTR';
            const fixedDecimals = isVTR ? 4 : 2;
            const trendEmoji = difference >= 0 ? '🟢' : '🔴';
            const sign = difference >= 0 ? '+' : '';
            const mainValueStr = val2.toLocaleString(undefined, { minimumFractionDigits: fixedDecimals, maximumFractionDigits: fixedDecimals });
            const diffStr = `${sign}${difference.toLocaleString(undefined, { minimumFractionDigits: fixedDecimals, maximumFractionDigits: fixedDecimals })}`;
            const changeStr = isFinite(percentageChange) ? `${sign}${percentageChange.toFixed(2)}%` : 'N/A';
            const emoji = getEmojiForMetric(metric);
            const namePart = `${emoji} ${metric}`.padEnd(22);
            const valuePart = mainValueStr.padStart(15);
            const changePart = `(${diffStr} | ${changeStr})`;
            const outputLine = `  ${namePart}: ${valuePart} ${trendEmoji} ${changePart}`;
            output.push(outputLine);
        });
    });
    return output.join('\n');
}

function generateHtmlReport(results1, results2, date1, date2, config, fileInfo) {
    const allGroupKeys = new Set([...Object.keys(results1), ...Object.keys(results2)]);
    let groupHtml = '';
    allGroupKeys.forEach(groupKey => {
        let metricsHtml = '';
        const metricsToCompare = [...config.metrics];
        if (config.vtrConfig) metricsToCompare.push('VTR');
        metricsToCompare.forEach(metric => {
            const val1 = (results1[groupKey] || {})[metric] || 0;
            const val2 = (results2[groupKey] || {})[metric] || 0;
            const difference = val2 - val1;
            const percentageChange = (val1 !== 0) ? (difference / val1) * 100 : Infinity;
            const changeClass = difference >= 0 ? 'positive' : 'negative';
            const fixedDecimals = metric === 'VTR' ? 4 : 2;
            metricsHtml += `
                <div class="metric-container">
                    <div class="metric-icon">${getIconForMetric(metric)}</div>
                    <div class="metric-content">
                        <div class="metric-title">${metric}</div>
                        <div class="metric-values">
                            <span class="value-pair"><strong>${date1}:</strong> ${val1.toFixed(fixedDecimals)}</span>
                            <span class="value-pair"><strong>${date2}:</strong> ${val2.toFixed(fixedDecimals)}</span>
                        </div>
                        <div class="metric-change">
                            <span class="${changeClass}"><strong>Diff:</strong> ${difference.toFixed(fixedDecimals)}</span>
                            <span class="${changeClass}"><strong>Change:</strong> ${isFinite(percentageChange) ? percentageChange.toFixed(2) + '%' : 'N/A'}</span>
                        </div>
                    </div>
                </div>
            `;
        });
        groupHtml += `
            <div class="group-container">
                <h2>${config.groupingColumn}: ${groupKey}</h2>
                <div class="metrics-grid">${metricsHtml}</div>
            </div>
        `;
    });
    return `
        <!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Comparative Analysis Report</title>
        <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f8f9fa; color: #212529; margin: 0; padding: 2rem; }
            .report-container { max-width: 1200px; margin: auto; }
            .report-header { background-color: #ffffff; border: 1px solid #dee2e6; border-radius: 8px; padding: 1.5rem; margin-bottom: 2rem; text-align: center; }
            .report-header h1 { margin: 0; color: #00529B; } .report-header p { margin: 0.5rem 0 0; color: #495057; font-size: 1.1rem; }
            .group-container { background-color: #ffffff; border: 1px solid #dee2e6; border-radius: 8px; margin-bottom: 2rem; padding: 1.5rem; }
            .group-container h2 { margin-top: 0; border-bottom: 2px solid #e9ecef; padding-bottom: 0.75rem; color: #343a40; }
            .metrics-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1.5rem; }
            .metric-container { display: flex; align-items: flex-start; background-color: #f8f9fa; padding: 1rem; border-radius: 6px; border: 1px solid #e9ecef; }
            .metric-icon { width: 40px; height: 40px; margin-right: 1rem; color: #007bff; }
            .metric-content { flex: 1; } .metric-title { font-weight: 600; font-size: 1.1rem; color: #212529; margin-bottom: 0.5rem; }
            .metric-values { display: flex; justify-content: space-between; gap: 1rem; margin-bottom: 0.5rem; } .value-pair { font-size: 1rem; }
            .metric-change { display: flex; justify-content: space-between; gap: 1rem; font-size: 0.9rem; color: #6c757d; }
            .metric-change .positive { color: #28a745; } .metric-change .negative { color: #dc3545; }
        </style></head><body>
            <div class="report-container">
                <div class="report-header"><h1>Comparative Analysis: ${config.displayName}</h1><p>${fileInfo}</p></div>
                ${groupHtml || '<h2>No data found to compare.</h2>'}
            </div>
        </body></html>
    `;
}

// ====================================================================================
// MAIN EXPORTED FUNCTION
// ====================================================================================

function runAnalysis(options) {
    const { isSingleFileMode, isWeeklyMode, isConciseMode, filePaths, config } = options;
    let results1, results2, date1, date2, fileInfo;

    try {
        if (isSingleFileMode) {
            if (isWeeklyMode) {
                const weeklyAnalysis = analyzeWeeklySingleFile(filePaths[0], config);
                ({ results1, results2, date1, date2, fileInfo } = weeklyAnalysis);
            } else { // Single file, Daily
                const dailyData = analyzeFile(filePaths[0], config, true);
                const dayCount = dailyData ? Object.keys(dailyData).length : 0;
                if (!dailyData || dayCount < 2) {
                     throw new Error(`Daily comparison failed. Only found ${dayCount} day(s) of data.`);
                }
                const dateConverter = (dateStr) => new Date(dateStr.split('/').reverse().join('-'));
                const sortedDates = Object.keys(dailyData).sort((a, b) => dateConverter(b) - dateConverter(a));
                date2 = sortedDates[0]; date1 = sortedDates[1];
                results1 = dailyData[date1]; results2 = dailyData[date2];
                fileInfo = `Comparing days ${date1} and ${date2} within ${path.basename(filePaths[0])}`;
            }
        } else { // Two-file mode
            if (!filePaths || filePaths.length < 2) throw new Error("Two files are required for this mode.");
            const analysis1 = analyzeFile(filePaths[0], config, false);
            const analysis2 = analyzeFile(filePaths[1], config, false);
            const dateObj1 = new Date(analysis1.fileDate.split('/').reverse().join('-'));
            const dateObj2 = new Date(analysis2.fileDate.split('/').reverse().join('-'));
            if (dateObj1 < dateObj2) {
                ({ groupedData: results1, fileDate: date1 } = analysis1);
                ({ groupedData: results2, fileDate: date2 } = analysis2);
                fileInfo = `Comparing ${path.basename(filePaths[0])} (${date1}) and ${path.basename(filePaths[1])} (${date2})`;
            } else {
                ({ groupedData: results1, fileDate: date1 } = analysis2);
                ({ groupedData: results2, fileDate: date2 } = analysis1);
                fileInfo = `Comparing ${path.basename(filePaths[1])} (${date1}) and ${path.basename(filePaths[0])} (${date2})`;
            }
        }

        if (isConciseMode) {
            return generateConciseTextReport(results1, results2, date1, date2, config);
        } else {
            return generateHtmlReport(results1, results2, date1, date2, config, fileInfo);
        }
    } catch (error) {
        console.error("Caught analysis error in runAnalysis:", error.message);
        throw error;
    }
}

module.exports = { runAnalysis, FLOW_A_CONFIG, FLOW_B_CONFIG };