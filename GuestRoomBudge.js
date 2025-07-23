// GuestRoomBudge.gs

/**
 * Guest Room Pricing Strategy Helper
 * Add this to your Google Apps Script for dynamic pricing analysis
 */

/**
 * Analyze and suggest optimal guest room pricing
 */
function analyzeGuestRoomPricing() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get current guest room data
  const guestRoomsSheet = ss.getSheetByName(GUEST_ROOMS_SHEET_NAME);
  const bookingsSheet = ss.getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  
  if (!guestRoomsSheet || !bookingsSheet) {
    ui.alert('Guest room sheets not found. Please initialize sheets first.');
    return;
  }
  
  // Analyze historical bookings
  const analysis = analyzeBookingHistory();
  
  // Calculate pricing recommendations
  const recommendations = calculatePricingRecommendations(analysis);
  
  // Display results
  showPricingAnalysis(analysis, recommendations);
}

/**
 * Analyze booking history for patterns
 */
function analyzeBookingHistory() {
  const bookingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  if (!bookingsSheet || bookingsSheet.getLastRow() < 2) {
    return {
      totalBookings: 0,
      averageStay: 0,
      occupancyRate: 0,
      averageRate: 0,
      weekdayOccupancy: 0,
      weekendOccupancy: 0,
      seasonalData: {}
    };
  }
  
  const data = bookingsSheet.getRange(2, 1, bookingsSheet.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
  
  let totalNights = 0;
  let totalRevenue = 0;
  let weekdayNights = 0;
  let weekendNights = 0;
  let bookingCount = 0;
  const monthlyData = {};
  
  const today = new Date();
  const oneYearAgo = new Date(today);
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
  
  data.forEach(booking => {
    const checkIn = new Date(booking[5]);
    const checkOut = new Date(booking[6]);
    const amount = booking[9];
    const status = booking[12];
    
    if (checkIn >= oneYearAgo && (status === 'Checked Out' || status === 'Checked In')) {
      bookingCount++;
      const nights = Math.ceil((checkOut - checkIn) / (1000 * 60 * 60 * 24));
      totalNights += nights;
      totalRevenue += amount;
      
      // Count weekday vs weekend nights
      for (let d = new Date(checkIn); d < checkOut; d.setDate(d.getDate() + 1)) {
        const dayOfWeek = d.getDay();
        if (dayOfWeek === 0 || dayOfWeek === 6) {
          weekendNights++;
        } else {
          weekdayNights++;
        }
        
        // Track monthly data
        const monthKey = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
        monthlyData[monthKey] = (monthlyData[monthKey] || 0) + 1;
      }
    }
  });
  
  // Calculate occupancy rates
  const totalPossibleNights = 365 * 2; // 2 guest rooms for 1 year
  const occupancyRate = totalNights / totalPossibleNights * 100;
  const weekdayOccupancy = weekdayNights / (260 * 2) * 100; // ~260 weekdays per year
  const weekendOccupancy = weekendNights / (105 * 2) * 100; // ~105 weekend days per year
  
  return {
    totalBookings: bookingCount,
    averageStay: bookingCount > 0 ? totalNights / bookingCount : 0,
    occupancyRate: occupancyRate.toFixed(1),
    averageRate: totalNights > 0 ? totalRevenue / totalNights : 0,
    weekdayOccupancy: weekdayOccupancy.toFixed(1),
    weekendOccupancy: weekendOccupancy.toFixed(1),
    seasonalData: monthlyData,
    totalRevenue: totalRevenue,
    totalNights: totalNights
  };
}

/**
 * Calculate pricing recommendations based on analysis
 */
function calculatePricingRecommendations(analysis) {
  const currentRooms = getCurrentGuestRoomRates();
  const recommendations = {
    weekdayRate: 0,
    weekendRate: 0,
    weeklyRate: 0,
    monthlyRate: 0,
    seasonalAdjustments: {},
    reasoning: []
  };
  
  // Base recommendations on occupancy
  if (analysis.occupancyRate > 80) {
    // High occupancy - increase prices
    recommendations.weekdayRate = currentRooms.avgDaily * 1.15;
    recommendations.weekendRate = currentRooms.avgDaily * 1.30;
    recommendations.reasoning.push('High occupancy (>80%) suggests room for price increase');
  } else if (analysis.occupancyRate > 60) {
    // Good occupancy - moderate pricing
    recommendations.weekdayRate = currentRooms.avgDaily * 1.05;
    recommendations.weekendRate = currentRooms.avgDaily * 1.20;
    recommendations.reasoning.push('Good occupancy (60-80%) supports current pricing with weekend premium');
  } else {
    // Low occupancy - competitive pricing
    recommendations.weekdayRate = currentRooms.avgDaily * 0.90;
    recommendations.weekendRate = currentRooms.avgDaily * 1.00;
    recommendations.reasoning.push('Lower occupancy (<60%) suggests more competitive pricing needed');
  }
  
  // Weekly and monthly rates
  recommendations.weeklyRate = recommendations.weekdayRate * 6.5; // ~7% discount
  recommendations.monthlyRate = recommendations.weekdayRate * 25; // ~17% discount
  
  // Seasonal adjustments
  const monthlyOccupancy = calculateMonthlyOccupancy(analysis.seasonalData);
  Object.entries(monthlyOccupancy).forEach(([month, occupancy]) => {
    if (occupancy > 80) {
      recommendations.seasonalAdjustments[month] = '+15%';
    } else if (occupancy < 40) {
      recommendations.seasonalAdjustments[month] = '-10%';
    }
  });
  
  // Weekend vs weekday analysis
  if (analysis.weekendOccupancy > analysis.weekdayOccupancy * 1.2) {
    recommendations.reasoning.push('Weekend demand is significantly higher - consider larger weekend premium');
  }
  
  return recommendations;
}

/**
 * Get current guest room rates
 */
function getCurrentGuestRoomRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_ROOMS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { avgDaily: 70, avgWeekly: 420 };
  }
  
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2).getValues();
  let totalDaily = 0;
  let totalWeekly = 0;
  let count = 0;
  
  data.forEach(room => {
    if (room[0]) {
      totalDaily += room[0];
      totalWeekly += room[1] || (room[0] * 7);
      count++;
    }
  });
  
  return {
    avgDaily: count > 0 ? totalDaily / count : 70,
    avgWeekly: count > 0 ? totalWeekly / count : 420
  };
}

/**
 * Calculate monthly occupancy percentages
 */
function calculateMonthlyOccupancy(seasonalData) {
  const monthlyOccupancy = {};
  const daysInMonth = {
    '01': 31, '02': 28, '03': 31, '04': 30,
    '05': 31, '06': 30, '07': 31, '08': 31,
    '09': 30, '10': 31, '11': 30, '12': 31
  };
  
  Object.entries(seasonalData).forEach(([monthKey, nights]) => {
    const month = monthKey.split('-')[1];
    const possibleNights = daysInMonth[month] * 2; // 2 rooms
    monthlyOccupancy[monthKey] = (nights / possibleNights * 100).toFixed(1);
  });
  
  return monthlyOccupancy;
}

/**
 * Display pricing analysis results
 */
function showPricingAnalysis(analysis, recommendations) {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>Guest Room Pricing Analysis</h3>
      
      <h4>Current Performance (Last 12 Months)</h4>
      <table style="border-collapse: collapse; margin-bottom: 20px;">
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Total Bookings:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">${analysis.totalBookings}</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Average Stay:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">${analysis.averageStay.toFixed(1)} nights</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Overall Occupancy:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">${analysis.occupancyRate}%</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Weekday Occupancy:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">${analysis.weekdayOccupancy}%</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Weekend Occupancy:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">${analysis.weekendOccupancy}%</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Average Daily Rate:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${analysis.averageRate.toFixed(2)}</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Total Revenue:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${analysis.totalRevenue.toFixed(2)}</td>
        </tr>
      </table>
      
      <h4>Pricing Recommendations</h4>
      <table style="border-collapse: collapse; margin-bottom: 20px;">
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Weekday Rate:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${recommendations.weekdayRate.toFixed(2)}/night</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Weekend Rate:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${recommendations.weekendRate.toFixed(2)}/night</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Weekly Rate:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${recommendations.weeklyRate.toFixed(2)} (7 nights)</td>
        </tr>
        <tr>
          <td style="padding: 5px; border: 1px solid #ddd;"><strong>Monthly Rate:</strong></td>
          <td style="padding: 5px; border: 1px solid #ddd;">$${recommendations.monthlyRate.toFixed(2)} (30 nights)</td>
        </tr>
      </table>
      
      <h4>Analysis & Reasoning</h4>
      <ul>
        ${recommendations.reasoning.map(r => `<li>${r}</li>`).join('')}
      </ul>
      
      <h4>Revenue Projection</h4>
      <p>If you implement recommended pricing with current occupancy levels:</p>
      <ul>
        <li>Projected Annual Revenue: $${calculateProjectedRevenue(analysis, recommendations).toFixed(2)}</li>
        <li>Revenue Increase: $${(calculateProjectedRevenue(analysis, recommendations) - analysis.totalRevenue).toFixed(2)}</li>
      </ul>
      
      <h4>Implementation Tips</h4>
      <ul>
        <li>Test new rates gradually - start with one room</li>
        <li>Monitor booking patterns after rate changes</li>
        <li>Consider seasonal events in your area</li>
        <li>Offer package deals for extended stays</li>
        <li>Use weekend rates for Friday & Saturday nights</li>
      </ul>
    </div>
  `;
  
  const htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(600)
      .setHeight(600);
  
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Guest Room Pricing Analysis');
}

/**
 * Calculate projected revenue with new rates
 */
function calculateProjectedRevenue(analysis, recommendations) {
  const weekdayNights = analysis.totalNights * (analysis.weekdayOccupancy / 100);
  const weekendNights = analysis.totalNights * (analysis.weekendOccupancy / 100);
  
  return (weekdayNights * recommendations.weekdayRate) + 
         (weekendNights * recommendations.weekendRate);
}

/**
 * Add pricing analysis to the menu
 */
function addPricingAnalysisToMenu() {
  // Add this item to your Guest Room Management submenu:
  // .addItem('Analyze Pricing Strategy', 'analyzeGuestRoomPricing')
}

/**
 * Create a dynamic pricing rule system
 */
function setupDynamicPricing() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Dynamic Pricing Setup',
    'This will create pricing rules based on:\n' +
    '• Occupancy levels\n' +
    '• Day of week\n' +
    '• Seasonal demand\n' +
    '• Length of stay\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Create a new sheet for pricing rules
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let pricingSheet = ss.getSheetByName('Pricing Rules');
  
  if (!pricingSheet) {
    pricingSheet = ss.insertSheet('Pricing Rules');
  } else {
    pricingSheet.clear();
  }
  
  // Set up pricing rules structure
  const headers = [
    'Rule Name',
    'Rule Type',
    'Condition',
    'Price Adjustment',
    'Priority',
    'Active',
    'Notes'
  ];
  
  const sampleRules = [
    ['Weekend Premium', 'Day of Week', 'Friday, Saturday', '+20%', 1, 'Yes', 'Applied to weekend nights'],
    ['Weekly Discount', 'Length of Stay', '7+ nights', '-10%', 2, 'Yes', 'Discount for week-long stays'],
    ['Monthly Discount', 'Length of Stay', '28+ nights', '-20%', 3, 'Yes', 'Discount for monthly stays'],
    ['High Season', 'Date Range', 'June-August', '+15%', 4, 'Yes', 'Summer peak pricing'],
    ['Low Season', 'Date Range', 'January-February', '-10%', 5, 'Yes', 'Winter discount'],
    ['Last Minute', 'Booking Window', 'Same day', '-15%', 6, 'Yes', 'Fill empty rooms'],
    ['Advance Booking', 'Booking Window', '30+ days', '-5%', 7, 'Yes', 'Reward early bookings']
  ];
  
  // Apply to sheet
  pricingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  pricingSheet.getRange(2, 1, sampleRules.length, sampleRules[0].length).setValues(sampleRules);
  
  // Format
  pricingSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  pricingSheet.setFrozenRows(1);
  
  for (let i = 1; i <= headers.length; i++) {
    pricingSheet.autoResizeColumn(i);
  }
  
  ui.alert('Dynamic Pricing Rules Created', 
    'Pricing rules have been set up in the "Pricing Rules" sheet.\n' +
    'You can modify these rules to match your pricing strategy.',
    ui.ButtonSet.OK);
}

/**
 * Apply dynamic pricing to a booking
 */
function calculateDynamicPrice(checkIn, checkOut, baseRate) {
  const pricingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pricing Rules');
  if (!pricingSheet || pricingSheet.getLastRow() < 2) {
    return baseRate;
  }
  
  const rules = pricingSheet.getRange(2, 1, pricingSheet.getLastRow() - 1, 6).getValues();
  let finalRate = baseRate;
  const appliedAdjustments = [];
  
  // Calculate stay details
  const nights = Math.ceil((checkOut - checkIn) / (1000 * 60 * 60 * 24));
  const dayOfWeek = checkIn.getDay();
  const bookingWindow = Math.ceil((checkIn - new Date()) / (1000 * 60 * 60 * 24));
  
  // Apply active rules by priority
  rules.sort((a, b) => a[4] - b[4]); // Sort by priority
  
  rules.forEach(rule => {
    if (rule[5] !== 'Yes') return; // Skip inactive rules
    
    const [name, type, condition, adjustment, priority, active] = rule;
    let applies = false;
    
    switch (type) {
      case 'Day of Week':
        const days = condition.split(',').map(d => d.trim());
        const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        if (days.includes(dayNames[dayOfWeek])) applies = true;
        break;
        
      case 'Length of Stay':
        const minNights = parseInt(condition.match(/\d+/)[0]);
        if (nights >= minNights) applies = true;
        break;
        
      case 'Booking Window':
        if (condition.includes('Same day') && bookingWindow === 0) applies = true;
        if (condition.includes('30+ days') && bookingWindow >= 30) applies = true;
        break;
        
      case 'Date Range':
        // Simplified - would need more complex date parsing in production
        const month = checkIn.getMonth();
        if (condition.includes('June-August') && month >= 5 && month <= 7) applies = true;
        if (condition.includes('January-February') && month <= 1) applies = true;
        break;
    }
    
    if (applies) {
      const percentMatch = adjustment.match(/([+-]\d+)%/);
      if (percentMatch) {
        const percent = parseInt(percentMatch[1]);
        finalRate = finalRate * (1 + percent / 100);
        appliedAdjustments.push(`${name}: ${adjustment}`);
      }
    }
  });
  
  return {
    rate: finalRate,
    adjustments: appliedAdjustments
  };
}
