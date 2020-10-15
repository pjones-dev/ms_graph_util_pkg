create or replace package ms_graph_util_pkg
as
 
  /*
 
  Purpose:      Package handles Microsoft Graph API
 
  Who     Date        Description
  ------  ----------  --------------------------------
  PJ      14.10.2020  Created
 
  */

  ------------
  -- types
  ------------  
   
  type t_event is record (
    --See https://docs.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0
    allowNewTimeProposals varchar2(10),
    attendees varchar2(32767),
    body clob,
    bodyPreview clob,
    categories varchar2(32767),
    event_id varchar2(2000),
    change_key varchar2(2000),
    createdDateTime date,
    end_date date,
    hasAttachments varchar2(10),
    iCalUId varchar2(2000),
    importance varchar2(10),
    isAllDay varchar2(10),
    isCancelled varchar2(10),
    isOnlineMeeting varchar2(10),
    isOrganizer varchar2(10),
    isReminderOn varchar2(10),
    lastModifiedDateTime date,
    location varchar2(2000),
    --locations clob, --TODO Add Support for multiple locations
    --onlineMeeting clob, --TODO Add Support for onlineMeeting details
    onlineMeetingProvider varchar2(20),
    onlineMeetingUrl varchar2(2000),
    organizer clob,
    originalEndTimeZone varchar2(2000),
    originalStart date,
    originalStartTimeZone varchar2(2000),
    --recurrence clob, --TODO Add Support for recurrence
    reminderMinutesBeforeStart number,
    responseRequested varchar2(10),
    responseStatus clob,
    sensitivity varchar2(20),
    seriesMasterId varchar2(2000),
    showAs varchar2(20),
    start_date date,
    subject varchar2(2000),
    transactionId varchar2(2000),
    event_type varchar2(20),
    webLink varchar2(2000),
    has_attachments varchar2(10)
    --attachments --TODO Add Support for attachments [ { "@odata.type": "microsoft.graph.attachment" } ],
    --calendar --TODO Add Support for graph calendar { "@odata.type": "microsoft.graph.calendar" },
    --extensions --TODO Add Support for extensions [ { "@odata.type": "microsoft.graph.extension" } ],
    --instances --TODO Add Support for instances [ { "@odata.type": "microsoft.graph.event" }],
    --multiValueExtendedProperties --TODO Add Support for multiValueExtendedProperties [ { "@odata.type": "microsoft.graph.multiValueLegacyExtendedProperty" }],
    --singleValueExtendedProperties --TODO Add Support for singleValueExtendedProperties [ { "@odata.type": "microsoft.graph.singleValueLegacyExtendedProperty" }]
  );
  
 
  ------------
  -- constants
  ------------
  
  --true and false
  g_true                        constant varchar2(10) := 'true';
  g_false                       constant varchar2(10) := 'false';

  --importance
  g_importance_low              constant varchar2(10) := 'low';
  g_importance_normal           constant varchar2(10) := 'normal';
  g_importance_high             constant varchar2(10) := 'high';

  --onlineMeetingProvider
  g_omp_teamsForBusiness        constant varchar2(20) := 'teamsForBusiness';
  g_omp_skypeForBusiness        constant varchar2(20) := 'skypeForBusiness';
  g_omp_skypeForConsumer        constant varchar2(20) := 'skypeForConsumer';

  --reminderMinutesBeforeStart
  g_default_rmbs                constant number := 10;

  --g_defaultTimeZone
  g_defaultTimeZone             constant varchar2(30) := 'GMT Standard Time';

  --sensitivity
  g_sensitivity_normal          constant varchar2(20) := 'normal';
  g_sensitivity_personal        constant varchar2(20) := 'personal';
  g_sensitivity_private         constant varchar2(20) := 'private';
  g_sensitivity_confidential    constant varchar2(20) := 'confidential';
  
  --showAs
  g_showAs_free                 constant varchar2(255) := 'free';
  g_showAs_tentative            constant varchar2(255) := 'tentative';
  g_showAs_busy                 constant varchar2(255) := 'busy';
  g_showAs_oof                  constant varchar2(255) := 'oof';
  g_showAs_workingElsewhere     constant varchar2(255) := 'workingElsewhere';
  g_showAs_unknown              constant varchar2(255) := 'unknown';

  --event type
  g_event_type_singleInstance   constant varchar2(255) := 'singleInstance';
  g_event_type_occurrence       constant varchar2(255) := 'occurrence';
  g_event_type_exception        constant varchar2(255) := 'exception';
  g_event_type_seriesMaster     constant varchar2(255) := 'seriesMaster';

  -----------------
  -- authentication
  -----------------
 
  -- initialize settings
  procedure init(p_url in varchar2,
                 p_login_url in varchar2,
                 p_oauth2_url in varchar2,
                 p_login_scope in varchar2,
                 p_tenantid in varchar2,
                 p_clientid in varchar2,
                 p_clientsecret in varchar2,
                 p_username in varchar2,
                 p_password in varchar2,
                 p_wallet_path in varchar2 := null,
                 p_wallet_password in varchar2 := null);
  
  -- get user access token
  function get_user_access_token return varchar2;
 
  -----------------
  -- calendar
  -----------------
 
  -- create calendar event
  function create_cal_event(p_event in t_event,
                           p_user in varchar2 := null,
                           p_calendar in varchar2 := null,
                           p_required_attendees in APEX_APPLICATION_GLOBAL.VC_ARR2,
                           p_categories in APEX_APPLICATION_GLOBAL.VC_ARR2) return varchar2;
 
  -- patch calendar event
  function patch_cal_event(p_event in t_event,
                           p_user in varchar2 := null,
                           p_calendar in varchar2 := null,
                           p_required_attendees in APEX_APPLICATION_GLOBAL.VC_ARR2,
                           p_categories in APEX_APPLICATION_GLOBAL.VC_ARR2) return CLOB;
  --delete calendar event
  procedure delete_cal_event(p_event_id in varchar2,
                             p_user in varchar2 := null,
                             p_calendar in varchar2 := null);

end ms_graph_util_pkg;
/

