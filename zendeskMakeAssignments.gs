/*
   For usage instructions, see README.md

   Parts of Code Copyright (c) 2013 by BetterCloud, Michael Stone
     All Rights Reserved, Not for distribution.
*/

function assignTickets() {
  const agentSheet = getAgentSheet();
  const logSheet = getTicketLogSheet();
  const debugSheet = getDebugSheet();

  const subdomain = getSubdomain();
  const username = getUsername();
  const token = getToken();

  // Get open tickets using Zendesk API
  const openTickets = fetchOpenTickets();

  // get the agents into an array
  var aAgentQueue = agentSheet.getRange("A2:G").getValues();

  var ticketId, tags, assigneeId, formType;

  debug('Number of open tickets: ' + openTickets.length);

  for(var i = 0; i < openTickets.length; i++) {
    ticketId = openTickets[i].id;
    tags = openTickets[i].tags.toString();
    assigneeId = openTickets[i].assignee_id;
    formType = parseFormType_(tags);

    // Update log sheet with ticket details
    if(isDebugMode()) {
      logSheet.insertRowBefore(2);
      logSheet.getRange("A2").setValue(Date());
      logSheet.getRange("B2").setValue(ticketId);
      // FIXME: Is this ever the case? We filter by assignee:none
      logSheet.getRange("C2").setValue(assigneeId || "Unassigned");
      logSheet.getRange("D2").setValue(tags);
      logSheet.getRange("E2").setValue(formType || "Invalid Form Type");
    }

    // FIXME: Is this ever the case? We filter by assignee:none
    // Ticket is already assigned, don't assign this ticket.
    if (assigneeId) { continue }
    // Skip ticket is formType is not available
    if(!formType) { continue }

    // get the currently available agent in the queue
    var agentAvailableItemNumber = seekNextAvailableAgentItem_(formType);

    // No agents available. Exit script.
    if(!agentAvailableItemNumber) { return }

    // FIXME: What are these magic numbers?
    var assigneeName = aAgentQueue[agentAvailableItemNumber][1];
    var agentUserID = aAgentQueue[agentAvailableItemNumber][2];
    var agentRowNumber = agentAvailableItemNumber + 2;

    // post the assignment
    if (postTicketAssignment_(subdomain, username, token, ticketId, agentUserID) == true) {

      // update log sheet
      logSheet.getRange("G2").setValue(assigneeName);
      logSheet.getRange("H2").setValue(agentUserID);

      // clear the previous assignment status
      agentSheet.getRange("D2:D").clearContent();

      // set the assignment status for the current agent
      agentSheet.getRange("D" + agentRowNumber).setValue("x");
      // agentSheet.getRange("H" + agentRowNumber).setValue(ticketId); // already provided in new log file

      Logger.log("Assigned ticket ID " + ticketId + " to support agent " + aAgentQueue[agentAvailableItemNumber][1] + ".");

      // ticket assigned, update log sheet
      logSheet.getRange("F2").setValue("Ticket Assigned Automatically");

    } else {
      // web service post failed, no action taken for this ticket, update log sheet
      logSheet.getRange("F2").setValue("Zendesk API Post Failed");
    }
  }
}

function setAssignedTicketsPerAgent() {
  const TICKETS_ASSIGNED_COLUMN = 'I';
  const AGENT_ID_COLUMN_INDEX = 2;
  const INDEX_OFFSET = 1;

  const agentSheet = getAgentSheet();
  const dataRange = agentSheet.getDataRange();
  const rowsWithData = dataRange.getValues();

  var agentId, ticketCount, rowIndex;

  rowsWithData.forEach(function(row, index) {
    // Skip header row
    if(index === 0) { return }

    rowIndex = index + INDEX_OFFSET;

    agentId = row[AGENT_ID_COLUMN_INDEX];

    ticketCount = fetchAssignedTickets(agentId);

    agentSheet.getRange(TICKETS_ASSIGNED_COLUMN + rowIndex).setValue(ticketCount);
  });
}

function fetchAssignedTickets(agentId) {
  const subdomain = getSubdomain();
  const username = getUsername();
  const token = getToken();
  const url = 'https://' + subdomain + '.zendesk.com/api/v2/search.json' + '?query=type:ticket status:open assignee:' + agentId;

  const authToken = username + "/token:" + token;
  const encodedAuthToken = Utilities.base64Encode(authToken);
  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Basic " + encodedAuthToken
    }
  };

  const jsonResponse = fetchJson(url, options);

  return jsonResponse.count;
}

function fakeOpenTickets() {
  const fakeTickets = [
    {
      id: 1,
      tags: [
        "am_notified",
        "customer_support",
        "greenhouse_recruiting_product",
      ],
      assignee_id: null,
    },
    {
      id: 2,
      tags: [
        "am_notified",
        "customer_support",
        "greenhouse_recruiting_product",
      ],
      assignee_id: null,
    },
    {
      id: 3,
      tags: [
        "am_notified",
        "customer_support",
        "push_to_veteran",
      ],
      assignee_id: null,
    },
    {
      id: 4,
      tags: [
        "am_notified",
        "customer_support",
        "push_to_veteran",
      ],
      assignee_id: null,
    },
    {
      id: 5,
      tags: [
        "am_notified",
        "customer_support",
        "draw_an_ace",
      ],
      assignee_id: null,
    },
    {
      id: 6,
      tags: [
        "am_notified",
        "customer_support",
        "greenhouse_recruiting_product",
      ],
      assignee_id: null,
    }
  ];

  return fakeTickets;
}

function fetchOpenTickets() {
  const subdomain = getSubdomain();
  const username = getUsername();
  const token = getToken();
  // Add additional filters to reduce the result set and execution time
  // by including tags and ordering by oldest first
  const searchUrl = "https://" + subdomain + ".zendesk.com/api/v2/search.json?" +
    "query=type:ticket status:new assignee:none tags:draw_an_ace order_by:created sort:asc";
  const authToken = username + "/token:" + token;
  const encodedAuthToken = Utilities.base64Encode(authToken);
  const options = {
    "method" : "get",
    "headers" : {
      "Content-type":"application/xml",
      "Authorization":  "Basic " + encodedAuthToken
    }
  };

  const jsonResponse = fetchJson(searchUrl, options);
  const MAX_TICKETS_PER_ROUND = getMaxTicketsPerRound();
  const maxOpenTickets = jsonResponse.results.slice(0, MAX_TICKETS_PER_ROUND);

  getDebugSheet().getRange("A2").setValue(maxOpenTickets);

  return maxOpenTickets;
}

function postTicketAssignment_(subdomain, username, token, ticketId, agentUserId) {
  token = username + "/token:" + token;
  var encode = Utilities.base64Encode(token);

  const payload = {
    "ticket": {
      "assignee_id" : parseInt(agentUserId)
    }
   };

   const options = {
      "method" : "put",
      "contentType" : "application/json",
      "headers" :
      {
        "Authorization" :  "Basic " + encode
      },
      "payload" : JSON.stringify(payload)
    };

  const url = "https://" + subdomain + ".zendesk.com/api/v2/tickets/" + ticketId + ".json";

  if(isReadonly()) {
    var assigneeId = agentUserId;
  } else {
    var ticket = fetchJson(url, options)
    var assigneeId = ticket.ticket.assignee_id.toString();
  }

  // posted successfully?
  return assigneeId == agentUserId;
}

function getSortedAgents() {
  var agents = getAgentSheet().getDataRange().getValues();

  // Remove the header row
  agents.shift();

  var sortedAgents = agents.map(function(agent, index) {
    return {
      row: index,
      data: agent,
    };
  });

  sortedAgents = sortedAgents.sort(sortByNumTicketsAssigned);

  logObject('sorted agents ===', sortedAgents.map(function(obj) { return obj.data; }));

  return sortedAgents;
}

function sortByNumTicketsAssigned(agentOne, agentTwo) {
  const TICKETS_ASSIGNED_INDEX = 8;

  return agentOne.data[TICKETS_ASSIGNED_INDEX] - agentTwo.data[TICKETS_ASSIGNED_INDEX];
}

function getAgentToAssign(ticket) {
  const sortedAgents = getSortedAgents();
  const ACTIVE_INDEX = 0;
  const TICKETS_ASSIGNED_INDEX = 8;

  var agent, numTicketsAssigned;

  for(var i = 0; i < sortedAgents.length; i++) {
    agent = sortedAgents[i].data;
    numTicketsAssigned = parseInt(agent[TICKETS_ASSIGNED_INDEX], 10);

    logObject('Agent ====', agent);

    if('Yes' === agent[ACTIVE_INDEX] && ticketTagsMatchAgentTags(ticket, agent) && numTicketsAssigned < getMaxTicketsPerAgent()) {
      return sortedAgents[i];
    }
  }

  return null;
}

function logObject(preface, obj) {
  var str = preface;

  str += ' {';

  for (key in Object.keys(obj)) {
    str += key + ':' + obj[key] + ', ';
  }

  str += ' }';

  debug(str);
}

function ticketTagsMatchAgentTags(ticket, agent) {
  const ticketTags = ticket.tags;
  const agentTags = tagsForAgent(agent);

  // Any agent can handle an untagged ticket
  return agentSatisifiesTicketReqs(agentTags, ticketTags);
}

function hasDrawAnAce(ticketTags) {
  const DRAW_AN_ACE = 'draw_an_ace';
  return ticketTags.indexOf(DRAW_AN_ACE) >= 0;
}

function hasPushToVeteran(ticketTags) {
  const PUSH_TO_VETERAN = 'push_to_veteran';
  return ticketTags.indexOf(PUSH_TO_VETERAN) >= 0;
}

function hasTicketRequirements(ticketTags) {
  return hasPushToVeteran(ticketTags) || hasDrawAnAce(ticketTags);
}

function agentSatisifiesTicketReqs(agentTags, ticketTags) {
  const DRAW_AN_ACE = 'draw_an_ace';
  const PUSH_TO_VETERAN = 'push_to_veteran';

  var requiredTags = [];
  if (hasDrawAnAce(ticketTags)) {
    requiredTags.push(DRAW_AN_ACE)
  }

  if (hasPushToVeteran(ticketTags)) {
    requiredTags.push(PUSH_TO_VETERAN)
  }

  if (requiredTags.length == 0) { return true; }

  for (var i = 0; i < requiredTags.length; i++) {
    if (agentTags.indexOf(requiredTags[i]) < 0) {
      return false
    }
  }

  return true
}

function tagsForAgent(agent) {
  const TEST_TAG_INDEX = 4;
  const DRAW_AN_ACE_TAG_INDEX = 5;
  const PUSH_TO_VETERAN_TAG_INDEX = 6;

  var agentTags = [];

  if(agent[TEST_TAG_INDEX].length > 0) {
    agentTags.push('rlbaker_test');
  }

  if(agent[DRAW_AN_ACE_TAG_INDEX].length > 0) {
    agentTags.push('draw_an_ace');
  }

  if(agent[PUSH_TO_VETERAN_TAG_INDEX].length > 0) {
    agentTags.push('push_to_veteran');
  }

  return agentTags;
}

function v2AssignTickets() {
  const AGENT_NAME_INDEX = 1;
  const AGENT_ID_INDEX = 2;
  const openTickets = fetchOpenTickets();

  const numTicketsToReview = Math.min(openTickets.length, getMaxTicketsPerRound());

  var ticket, agentToAssign, agentRow, agentData, agentId;

  debug("Number of tickets to assign:" + numTicketsToReview);

  for(var index = 0; index < numTicketsToReview; index++) {
    ticket = openTickets[index];
    debug('Tags for ticket ' + ticket.id + ': ' + ticket.tags);

    if(ticket.assignee_id) { debug('Skip ticket (already assigned): ' + ticket.id); continue; }

    agentToAssign = getAgentToAssign(ticket);

    if(agentToAssign) {
      agentData = agentToAssign.data
      agentId = agentData[AGENT_ID_INDEX];

      assignTicketInZendesk(ticket, agentId);
      debug('Assign ticket ' + ticket.id + ' to ' + agentData[AGENT_NAME_INDEX]);

      setAssignedTicketsPerAgent();
      logTicket(ticket, agentData)
    } else {
      debug('Skip ticket (no agents for ticket): ' + ticket.id);
    }
  }
}

function assignTicketInZendesk(ticket, agentId) {
  const url = "https://" + getSubdomain() + ".zendesk.com/api/v2/tickets/" + ticket.id + ".json";

  const payload = {
    "ticket": {
      "assignee_id" : parseInt(agentId)
    }
   };

  const options = {
    'method': 'put',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Basic ' + getAuthorizationToken(),
    },
    payload: JSON.stringify(payload),
  }

  if(isReadonly()) {
    debug('Fetch: ' + url + ' for agent ' + agentId);
  } else {
    fetchJson(url, options)
  }
}

function getAuthorizationToken() {
  const username = getUsername();
  const token = getToken();

  return Utilities.base64Encode(username + '/token:' + token);
}

function logTicket(ticket, agentData) {
  if(!isDebugMode()) { return }

  const logSheet = getTicketLogSheet();
  const agentId = agentData[2];
  const agentName = agentData[1];

  logSheet.insertRowBefore(2);
  logSheet.getRange("A2").setValue(Date());
  logSheet.getRange("B2").setValue(ticket.id);
  logSheet.getRange("C2").setValue(ticket.assignee_id || "Unassigned");
  logSheet.getRange("D2").setValue(ticket.tags.toString());
//  logSheet.getRange("E2").setValue(formType || "Invalid Form Type");
  logSheet.getRange("F2").setValue("Ticket Assigned Automatically");
  logSheet.getRange("G2").setValue(agentName);
  logSheet.getRange("H2").setValue(agentId);
}

function seekNextAvailableAgentItem_(formType) {
  const agentSheet = getAgentSheet();
  const formColumn = getFormColumn_(formType);

  // get the agents into an array
  // scoped to col J for a maxRange of 10, but needs to be updated if more than 12
  // FIXME: Column should reflect what is configured by maxTicketsPerAgent
  // Ideally, this would be dynamic
  const aAgentQueue = agentSheet.getRange("A2:J").getValues();
  const rowsWithData = agentSheet.getDataRange().getValues();

  // locate the previous agent assigned
  const previouslyAssignedAgentItem = seekPreviouslyAssignedAgentItem_(aAgentQueue);

  // if the previously assigned agent is in the last row, start at the top and search for someone other than the previously assigned agent
  if ((previouslyAssignedAgentItem == aAgentQueue.length) || (previouslyAssignedAgentItem == 0)) {
    // start at the top - find the next active agent
    for (var j = 0; j < aAgentQueue.length; j++) {
      if ((aAgentQueue[j][0] == "Yes") && (aAgentQueue[j][3] != "x") && (aAgentQueue[j][formColumn] == "x")) {
        return j;
      }
    }
    return(previouslyAssignedAgentItem);
  } else {
    // see if any agents below the previously assigned agent are able to handle this form type
    // search to the bottom of the list
    for(var j = (previouslyAssignedAgentItem + 1); j < aAgentQueue.length; j++) {
      if((aAgentQueue[j][0] == "Yes") && (aAgentQueue[j][3] != "x") && (aAgentQueue[j][formColumn] == "x")) {
        return j;
      }
    }
    // if not found, start over at the top of the list
    for(var j = 0; j < aAgentQueue.length; j++) {
      if ((aAgentQueue[j][0] == "Yes") && (aAgentQueue[j][3] != "x") && (aAgentQueue[j][formColumn] == "x")) {
        return j;
      }
    }
  }

  return null;
}

function seekPreviouslyAssignedAgentItem_(rows) {
  const QUEUE_COLUMN = 3;

  for(var i = 0; i < rows.length; i++) {
    if(rows[i][QUEUE_COLUMN] === "x") {
      return i;
    }
  }

  return 0;
}

// Match ticket type to the ones that different people can handle
// There is an x each tag column if the agent should be assigned tickets
// with that tag.
// Currently, tags are:
// draw_an_ace for gambler column
// rlbaker_test for test column
// push_to_veteran for legacy gambler column
function parseFormType_(tags) {
  const agentSheet = getAgentSheet();
  const maxRange = getMaxTicketsPerRound();

  // determine the dynamic range
  var dynamicRangeMax;
  for(var i = 5; i < maxRange; i++) {
    if(agentSheet.getRange(1, i, 1, 1).getNote() == "") {
      dynamicRangeMax = i - 5;
      i = maxRange;
    }
  }

  // get dyanamic range into array
  // Get comment on a specific cell. the comment is the tag.
  var aFormTypes = agentSheet.getRange(1, 5, 1, dynamicRangeMax).getNotes();
  debug("aFormTypes: " + aFormTypes.toString() + ", " + "Width: " + dynamicRangeMax);

  // look at each form type column
  for(var i = 0; i < dynamicRangeMax; i++) {
    if(tags.toString().indexOf(aFormTypes[0][i]) > -1) {
      debug("Found FormType Match!  Returning " + aFormTypes[0][i]);

      // Returns if there's an x for a specifc tag?
      return(aFormTypes[0][i]);
    }
  }

  return null;
}

function getFormColumn_(formType) {
  const maxRange = getMaxTicketsPerRound();
  const agentSheet = getAgentSheet();

  debug("getFormColumn_ maxRange:" + maxRange);

  // determine the dynamic range
  for (var i = 5; i < maxRange; i++) {
    debug("i:" + i + " :: " + agentSheet.getRange(1, i, 1, 1).getNote());
    if (agentSheet.getRange(1, i, 1, 1).getNote() == formType) {
      return(i-1);
    }
  }

  return(-1);
}

function isDebugMode() {
  return(PropertiesService.getScriptProperties().getProperty('debug') === 'true');
}

function debug(data) {
  if(isDebugMode()) {
    Logger.log(data);
  }
}

function getCurrentUtcHour() {
  return new Date().getUTCHours();
}

function getCurrentUtcDay() {
  return new Date().getUTCDay();
}

// Set if agent status to yes or no based on their working hours.
function setAgentStatuses() {
  const agentSheet = getSpreadsheet().getSheetByName('Support Agents');

  // Compensate for zero offset index
  const OFFSET = 1;

  const AGENT_NAME_INDEX = 1;
  const AGENT_WORKING_HOURS_INDEX = 7;

  const AGENT_STATUS_COLUMN = 'A';
  const AGENT_NAME_COLUMN = 'B';

  const dataRange = agentSheet.getDataRange();
  const rowsWithData = dataRange.getValues();

  var shiftData, shiftStart, shiftEnd;
  var agentName, agentActive;
  var rowIndex;

  rowsWithData.forEach(function(row, index) {
    // Skip header row
    if(index === 0) { return }

    rowIndex = index + OFFSET;

    agentName = row[AGENT_NAME_INDEX].trim();

    shiftData = row[AGENT_WORKING_HOURS_INDEX].split('-');
    shiftStart = shiftData[0];
    shiftEnd = shiftData[1];

    agentNameRange = agentSheet.getRange(AGENT_NAME_COLUMN + rowIndex);

    if(hasOverride(agentNameRange)) {
      agentActive = 'No';
    } else if(isAgentActive(shiftStart, shiftEnd)) {
      agentActive = 'Yes';
    } else {
      agentActive = 'No';
    }

    agentSheet.getRange(AGENT_STATUS_COLUMN + rowIndex).setValue(agentActive);

    debug('Set ' + agentName + ' (shift: ' + shiftData +  ') status to: ' + agentActive + ' for day: ' + getCurrentUtcDay() + ' and hour: ' + getCurrentUtcHour());
  });
}

function isAgentActive(startHour, endHour) {
  const weekday = isESTWeekDay();
  const workingHour = isWithinWorkingHours(startHour, endHour);

  debug('  is weekday: ' + weekday + ' working: ' + workingHour);

  return weekday && workingHour;
}

// We strike through the agent's name to override them as not active
function hasOverride(range) {
  return range.getFontLine() === 'line-through';
}

// startHour and endHour are string representations of integers from 0 to 23
function isWithinWorkingHours(startHour, endHour) {
  const currentHour = getCurrentUtcHour();
  startHour = parseInt(startHour, 10);
  endHour = parseInt(endHour, 10);

  if(isDaylightSavingsTime()) {
    startHour = startHour + 1
    endHour = endHour + 1
  }

  if(currentHour === startHour) {
    return true;
  } else if(currentHour === endHour) {
    return false;
  } else if(endHour < startHour) {
    // endHour is the next day
    endHour = endHour + 24

    if(currentHour === 0) {
      // startHour is the previous day
      startHour = startHour - 24
    }

    return (currentHour >= startHour && currentHour < endHour);
  } else {
    return (currentHour >= startHour && currentHour < endHour);
  }
}

function isDaylightSavingsTime() {
  return 'true' === PropertiesService.getScriptProperties().getProperty('isDaylightSavingsTime');
}

function isReadonly() {
  return 'true' === PropertiesService.getScriptProperties().getProperty('readonly');
}

function offsetFromUTC() {
  const EST_OFFSET = -5;

  var offset = EST_OFFSET;

  if(isDaylightSavingsTime()) {
    offset = offset + 1;
  }

  return offset;
}

function isESTWeekDay() {
  const FRIDAY = 5
  const SATURDAY = 6;
  const SUNDAY = 0;

  const now = new Date();
  const hour = getCurrentUtcHour();

  var day = getCurrentUtcDay();

  // Correct for ET day vs. UTC day
  if(hour + offsetFromUTC() < 0) {
    day = day - 1

    if(day < 0) {
      day = FRIDAY;
    }
  }

  return day !== SUNDAY && day !== SATURDAY;
}


function fetchJson(url, options) {
  debug("Fetch: " + url);

  const response = UrlFetchApp.fetch(url, options);

  return JSON.parse(response);
}

// Getters
function getSpreadsheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('sheetId') || "SHEET_ID";

  return SpreadsheetApp.openById(sheetId);
}

function getAgentSheet() {
  return getSpreadsheet().getSheetByName('Support Agents');
}

function getConfigurationSheet() {
  return getSpreadsheet().getSheetByName('Configuration');
}

function getDebugSheet() {
  return getSpreadsheet().getSheetByName('Debug Log');
}

function getTicketLogSheet() {
  return getSpreadsheet().getSheetByName('Assignment Log');
}

function getMaxTicketsPerRound() {
  const maxTicketsPerRound = PropertiesService.getScriptProperties().getProperty('maxTicketsPerRound');

  return parseInt(maxTicketsPerRound, 10);
}

function getMaxTicketsPerAgent() {
  const maxTicketsPerAgent = PropertiesService.getScriptProperties().getProperty('maxTicketsPerAgent');

  return parseInt(maxTicketsPerAgent, 10);
}

function getSubdomain() {
  return PropertiesService.getScriptProperties().getProperty('subdomain');
}

function getUsername() {
  return PropertiesService.getScriptProperties().getProperty('username');
}

function getToken() {
  return PropertiesService.getScriptProperties().getProperty('token');
}

function setConfiguration() {
  const spreadsheet = getSpreadsheet();

  const configurationSheet = spreadsheet.getSheetByName('Configuration');

  const subdomain = configurationSheet.getRange('B1').getValue();
  PropertiesService.getScriptProperties().setProperty('subdomain', subdomain);

  const username = configurationSheet.getRange('B2').getValue();
  PropertiesService.getScriptProperties().setProperty('username', username);

  const zendeskToken = configurationSheet.getRange('B3').getValue();
  PropertiesService.getScriptProperties().setProperty('token', zendeskToken);

  const isDebug = 'true';
  PropertiesService.getScriptProperties().setProperty('debug', isDebug);

  const isDaylightSavingsTime = configurationSheet.getRange('B5').getValue() === 'yes';
  PropertiesService.getScriptProperties().setProperty('isDaylightSavingsTime', isDaylightSavingsTime);

  const readonly = configurationSheet.getRange('B7').getValue();
  PropertiesService.getScriptProperties().setProperty('readonly', readonly);

  const maxTicketsPerRound = configurationSheet.getRange('B6').getValue();
  PropertiesService.getScriptProperties().setProperty('maxTicketsPerRound', maxTicketsPerRound);

  const maxTicketsPerAgent = configurationSheet.getRange('B8').getValue() || '5';
  PropertiesService.getScriptProperties().setProperty('maxTicketsPerAgent', maxTicketsPerAgent);

  const sheetId = configurationSheet.getRange('B9');
  if(sheetId && sheetId.length > 0) {
      PropertiesService.getScriptProperties().setProperty('sheetId', sheetId);
  }

  debug('Set Configuration:');
  debug(PropertiesService.getScriptProperties().getProperties());
}

function main() {
  debug('==== Configuration Values ===');
  setConfiguration();

  debug('==== Set Agent Statuses ===');
  setAgentStatuses();

  debug('==== Assign Tickets ===');
  v2AssignTickets();
  // assignTickets();
}
