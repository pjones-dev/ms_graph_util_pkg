create or replace package body ms_graph_util_pkg
as
 
  /*
 
  Purpose:      Package handles Microsoft Graph API
 
  Who     Date        Description
  ------  ----------  --------------------------------
  PJ      14.10.2020  Created
 
  */

  g_a_content_type_header_name  constant varchar2(30) := 'Content-Type';
  g_a_content_type_header_value constant varchar2(40) := 'application/x-www-form-urlencoded';

  g_authorization_header_name  constant varchar2(30) := 'Authorization';

  g_content_type_header_name  constant varchar2(30) := 'Content-Type';
  g_content_type_header_value constant varchar2(40) := 'application/json';

  g_get constant varchar2(30) := 'GET';
  g_post constant varchar2(30) := 'POST';
  g_patch constant varchar2(30) := 'PATCH';
  g_delete constant varchar2(30) := 'DELETE';

  g_url                        varchar2(2000);
  g_login_url                  varchar2(2000);
  g_oauth2_url                 varchar2(2000);
  g_login_scope                varchar2(2000);
  g_tenantid                   varchar2(2000);
  g_clientid                   varchar2(2000);
  g_clientsecret               varchar2(2000);
  g_username                   varchar2(2000);
  g_password                   varchar2(2000);
  g_wallet_path                varchar2(2000);
  g_wallet_password            varchar2(2000);

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
               p_wallet_password in varchar2 := null) 
as
begin
  
  g_url              := p_url;
  g_login_url        := p_login_url;
  g_oauth2_url       := p_oauth2_url;
  g_login_scope      := p_login_scope;
  g_tenantid         := p_tenantid;
  g_clientid         := p_clientid;
  g_clientsecret     := p_clientsecret;
  g_username         := p_username;
  g_password         := p_password;
  g_wallet_path      := p_wallet_path;
  g_wallet_password  := p_wallet_password;
 
end init;
 

procedure raise_error(p_error_message in varchar2)
as
begin

  raise_application_error (-20000, 'MS Graph API: ' || p_error_message);

end raise_error;


procedure assert_init 
as
begin
 
  if g_url is null then
    raise_error ('URL not specified. Please call .init() at least once per database session.');
  end if;

  if g_login_url is null then
    raise_error ('Login URL not specified. Please call .init() at least once per database session.');
  end if;

  if g_oauth2_url is null then
    raise_error ('OAUTH2 URL not specified. Please call .init() at least once per database session.');
  end if;

  if g_login_scope is null then
    raise_error ('Login Scope not specified. Please call .init() at least once per database session.');
  end if;

  if g_tenantid is null then
    raise_error ('Tenent ID not specified. Please call .init() at least once per database session.');
  end if;

  if g_clientid is null then
    raise_error ('Client ID not specified. Please call .init() at least once per database session.');
  end if;

  if g_clientsecret is null then
    raise_error ('Client Secret not specified. Please call .init() at least once per database session.');
  end if;

  if g_username is null then
    raise_error ('Username not specified. Please call .init() at least once per database session.');
  end if;

  if g_password is null then
    raise_error ('Password not specified. Please call .init() at least once per database session.');
  end if;

  if lower(g_url) like 'https%' then

    if g_wallet_path is null then
      raise_error ('Oracle Wallet path not specified. When using HTTPS you must reference an Oracle Wallet.');
    end if;

    if g_wallet_password is null then
      raise_error ('Oracle Wallet password not specified. When using HTTPS you must reference an Oracle Wallet.');
    end if;

  end if;
 
end assert_init;


function get_date(p_date_str in varchar2) return date
as
  l_returnvalue date;
begin
  
  begin
    l_returnvalue := to_date(substr(p_date_str,1,19), 'YYYY-MM-DD"T"HH24:MI:SS');
  exception
    when others then
      l_returnvalue := null;
  end;
  
  return l_returnvalue;

end get_date;


function get_date_str(p_date in date) return varchar2
as
  l_returnvalue varchar2(20);
begin
  
  l_returnvalue := to_char(p_date, 'YYYY-MM-DD"T"HH24:MI:SS');
  
  return l_returnvalue;

end get_date_str;

--check for errors
procedure check_for_errors(p_response_code in number,
                           p_response in clob)
as
  l_response_code varchar2(255);
  l_data          apex_json.t_values;
  l_count         PLS_INTEGER := 0;
  l_error         varchar2(255);
  l_error_desc    varchar2(4000);
  l_code          varchar2(255);
  l_message       varchar2(255);
begin
  
  l_response_code := p_response_code;
  
  if l_response_code not in (200,201,204) then
    
    if l_response_code = 400 then
      raise_error(l_response_code || ': ' || p_response);
    else 
    
      apex_json.parse(p_values => l_data,
                      p_source => p_response);
      
      
      l_error := apex_json.get_varchar2(p_path   => 'error',
                                        p_values => l_data); 

      l_error_desc := apex_json.get_varchar2(p_path   => 'error_description',
                                             p_values => l_data); 

      l_code := apex_json.get_varchar2(p_path   => 'error.code',
                                       p_values => l_data); 

      l_message := apex_json.get_varchar2(p_path   => 'error.message',
                                          p_values => l_data);

      raise_error (l_response_code || ': ' || p_response || ': ' || l_error || ': ' || l_error_desc || ': ' || l_code || ': ' || l_message);

    end if;

  end if;

end check_for_errors;

function get_user_access_token return varchar2
as
 l_response CLOB;
 l_data     apex_json.t_values;
 l_token_type VARCHAR2(50);
 l_access_token VARCHAR2(4000);
begin
  
  apex_web_service.g_request_headers.delete();
  apex_web_service.g_request_headers(1).name := g_a_content_type_header_name;
  apex_web_service.g_request_headers(1).value := g_a_content_type_header_value;

  l_response  := apex_web_service.make_rest_request(p_url         => g_login_url || g_tenantid || g_oauth2_url,
                                                    p_http_method => g_post,
                                                    p_parm_name   => apex_util.string_to_table('grant_type:client_id:client_secret:scope:userName:password'),
                                                    p_parm_value  => apex_util.string_to_table('password^'
                                                                                              || g_clientid || '^'
                                                                                              || g_clientsecret || '^'
                                                                                              || g_login_scope || '^'
                                                                                              || g_username || '^'
                                                                                              || g_password, '^'),
                                                    p_wallet_path => g_wallet_path,
                                                    p_wallet_pwd  => g_wallet_password);
  
  check_for_errors(apex_web_service.g_status_code,l_response);

  IF apex_web_service.g_status_code = 200 THEN
    apex_json.parse(p_values => l_data,
                    p_source => l_response);

    l_token_type          := apex_json.get_varchar2(p_path   => 'token_type',
                                                    p_values => l_data);

    l_access_token        := apex_json.get_varchar2(p_path   => 'access_token',
                                                    p_values => l_data);

    return l_token_type || ' ' || l_access_token;

  ELSE
    --check_for_errors should have picked up everything but if not log it
    raise_error('Other Error: ' || apex_web_service.g_status_code || ': ' || l_response);
  END IF;
end get_user_access_token;
 
--make rest request
function make_request(p_url in varchar2,
                      p_http_method in varchar2,
                      p_body in clob) return clob
as
  l_response CLOB;
  l_body CLOB  DEFAULT NVL(p_body,EMPTY_CLOB());
  l_token VARCHAR2(4000);
begin
  
  assert_init;
  
  l_token := get_user_access_token;
  
  apex_web_service.g_request_headers.delete();
  apex_web_service.g_request_headers(1).name := g_authorization_header_name;
  apex_web_service.g_request_headers(1).value := l_token;
  apex_web_service.g_request_headers(2).name := g_content_type_header_name;
  apex_web_service.g_request_headers(2).value := g_content_type_header_value;
  
  l_response  := apex_web_service.make_rest_request(p_url         => p_url,
                                                      p_http_method => p_http_method,
                                                      p_username    => null,
                                                      p_password    => null,
                                                      p_body        => l_body,
                                                      p_wallet_path => g_wallet_path,
                                                      p_wallet_pwd  => g_wallet_password);

  check_for_errors(apex_web_service.g_status_code,l_response);
  
  return l_response;

end make_request;

function create_cal_event(p_event in t_event,
                          p_user in varchar2 := null,
                          p_calendar in varchar2 := null,
                          p_required_attendees in APEX_APPLICATION_GLOBAL.VC_ARR2,
                          p_categories in APEX_APPLICATION_GLOBAL.VC_ARR2) return varchar2
as
  
  l_url                          varchar2(2000);
  l_http_method                  constant varchar2(30) := g_post;
  l_response                     CLOB;
  l_returnvalue                  varchar2(2000);
  l_data                         apex_json.t_values;
  
  
  function get_required_attendees return clob
  as
    l_returnvalue clob := '[]';
  begin

    if p_required_attendees.count > 0 then
    
      for i in p_required_attendees.first .. p_required_attendees.last loop
        if i = 1 THEN
          l_returnvalue := '{"emailAddress": {"address":"' || p_required_attendees(i) || '"},"type": "required"}';
        else
          l_returnvalue := l_returnvalue || ',{"emailAddress": {"address":"' || p_required_attendees(i) || '"},"type": "required"}';
        end if;
      end loop;
      
      l_returnvalue := '[' || l_returnvalue || ']';
      
    end if;
  
    return l_returnvalue;
  
  end get_required_attendees;
 
  function get_categories return clob
  as
    l_returnvalue clob := '[]';
  begin

    if p_categories.count > 0 then

      for i in p_categories.first .. p_categories.last loop
        if i = 1 THEN
          l_returnvalue := '"' || p_categories(i) || '"';
        else
          l_returnvalue := l_returnvalue || ',"' || p_categories(i) || '"';
        end if;
      end loop;

      l_returnvalue := '[' || l_returnvalue || ']';

    end if;

    return l_returnvalue;

  end get_categories;
  
  function get_request_body return clob
  as
    l_returnvalue clob;
  begin

    l_returnvalue := '{
                        "allowNewTimeProposals": ' ||  nvl(p_event.allowNewTimeProposals,g_false) || ',
                        "attendees": ' || get_required_attendees || ',
                        "body": {"contentType": "HTML","content": "' || APEX_ESCAPE.JSON(p_event.body) || '"},' ||
                        case when p_event.bodyPreview is not null then '"bodyPreview": "' || APEX_ESCAPE.JSON(p_event.bodyPreview) || '},' else '' end || '
                        "categories": ' || get_categories || ',
                        "end": {"dateTime": "' || get_date_str(nvl(p_event.end_date, p_event.start_date+1)) || '","timeZone": "' || nvl(p_event.originalEndTimeZone,g_defaultTimeZone) || '"},
                        "hasAttachments": ' || nvl(p_event.has_attachments,g_false) ||',
                        "importance": "' || nvl(p_event.importance,g_importance_normal) || '",
                        "isAllDay": ' || nvl(p_event.isAllDay,g_false) ||',
                        "isOnlineMeeting": ' || nvl(p_event.isOnlineMeeting,g_false) ||',
                        "isOrganizer": ' || nvl(p_event.isOrganizer,g_true) ||',
                        "isReminderOn": ' || nvl(p_event.isReminderOn,g_true) ||',' ||
                        case when p_event.location is not null then '"location": {"displayName": "' || APEX_ESCAPE.JSON(p_event.location) || '"},' else '' end || 
                        case when p_event.organizer is not null then '"organizer": {"emailAddress": "' || APEX_ESCAPE.JSON(p_event.organizer) || '"},' else '' end || 
                        case when p_event.onlineMeetingProvider is not null then '"onlineMeetingProvider": ' || p_event.onlineMeetingProvider || ',' else '' end || 
                        case when p_event.onlineMeetingProvider is not null and p_event.onlineMeetingUrl is not null then '"onlineMeetingUrl": "' || p_event.onlineMeetingProvider || '",' else '' end || 
                        case when nvl(p_event.isReminderOn,g_true) = 'true' then '"reminderMinutesBeforeStart": ' || nvl(p_event.reminderMinutesBeforeStart,g_default_rmbs) || ',' else '' end || '
                        "responseRequested": ' || nvl(p_event.responseRequested,g_true) || ',
                        "sensitivity": "' || nvl(p_event.responseRequested,g_sensitivity_normal) || '",
                        "showAs": "' || nvl(p_event.showAs,g_showAs_free) || '",
                        "start": {"dateTime": "' || get_date_str(p_event.start_date) || '","timeZone": "' || nvl(p_event.originalStartTimeZone,g_defaultTimeZone) || '"},
                        "subject": "' || nvl(APEX_ESCAPE.JSON(p_event.subject),'No Subject') || '",
                        "type": "' || nvl(p_event.event_type,g_event_type_singleInstance) || '"
                      }';

    return l_returnvalue;

  end get_request_body;

begin
  
  IF p_user IS NOT NULL and p_calendar IS NOT NULL THEN
    l_url := g_url || 'users/' || p_user || '/calendars/' || p_calendar;
  ELSE 
    l_url := g_url || 'me/calendar/events';
  END IF;

  l_response := make_request(p_url => l_url,
                             p_http_method => l_http_method,
                             p_body => get_request_body);

  IF apex_web_service.g_status_code = 201 THEN
    apex_json.parse(p_values => l_data,
                    p_source => l_response);

    l_returnvalue := apex_json.get_varchar2(p_path=>'id',
                                            p_values => l_data);

  ELSE 
    --check_for_errors should have picked up everything but if not log it
    raise_error('Other Error: ' || apex_web_service.g_status_code || ': ' || l_response);
  END IF;
  
  return l_returnvalue;

end create_cal_event;

function patch_cal_event(p_event in t_event,
                         p_user in varchar2 := null,
                         p_calendar in varchar2 := null,
                         p_required_attendees in APEX_APPLICATION_GLOBAL.VC_ARR2,
                         p_categories in APEX_APPLICATION_GLOBAL.VC_ARR2) return CLOB
as
  
  l_url                          varchar2(2000);
  l_http_method                  constant varchar2(30) := g_patch;
  l_response                     CLOB;
  l_returnvalue                  varchar2(2000);
  l_data                         apex_json.t_values;
  
  function get_required_attendees return clob
  as
    l_returnvalue clob := '[]';
  begin

    if p_required_attendees.count > 0 then
    
      for i in p_required_attendees.first .. p_required_attendees.last loop
        if i = 1 THEN
          l_returnvalue := '{"emailAddress": {"address":"' || p_required_attendees(i) || '"},"type": "required"}';
        else
          l_returnvalue := l_returnvalue || ',{"emailAddress": {"address":"' || p_required_attendees(i) || '"},"type": "required"}';
        end if;
      end loop;
      
      l_returnvalue := '[' || l_returnvalue || ']';
      
    end if;
  
    return l_returnvalue;
  
  end get_required_attendees;
 
  function get_categories return clob
  as
    l_returnvalue clob := '[]';
  begin

    if p_categories.count > 0 then

      for i in p_categories.first .. p_categories.last loop
        if i = 1 THEN
          l_returnvalue := '"' || p_categories(i) || '"';
        else
          l_returnvalue := l_returnvalue || ',"' || p_categories(i) || '"';
        end if;
      end loop;

      l_returnvalue := '[' || l_returnvalue || ']';

    end if;

    return l_returnvalue;

  end get_categories;
  
  function get_request_body return clob
  as
    l_returnvalue clob;
  begin

    l_returnvalue := '{' ||
                        case when p_event.allowNewTimeProposals is not null then '"allowNewTimeProposals": ' ||  nvl(p_event.allowNewTimeProposals,g_false) || ',' else '' end || 
                        case when p_required_attendees.count > 0 then '"attendees": ' || get_required_attendees || ',' else '' end || 
                        case when p_event.body is not null then '"body": {"contentType": "HTML","content": "' || APEX_ESCAPE.JSON(p_event.body) || '"},' else '' end || 
                        case when p_event.bodyPreview is not null then '"bodyPreview": "' || APEX_ESCAPE.JSON(p_event.bodyPreview) || '},' else '' end ||
                        case when p_categories.count > 0 then '"categories": ' || get_categories || ',' else '' end || 
                        case when p_event.end_date is not null then '"end": {"dateTime": "' || get_date_str(nvl(p_event.end_date, p_event.start_date+1)) || '","timeZone": "' || nvl(p_event.originalEndTimeZone,g_defaultTimeZone) || '"},' else '' end ||
                        case when p_event.hasAttachments is not null then '"hasAttachments": ' || nvl(p_event.has_attachments,g_false) ||',' else '' end || 
                        case when p_event.importance is not null then '"importance": "' || nvl(p_event.importance,g_importance_normal) || '",' else '' end || 
                        case when p_event.isAllDay is not null then '"isAllDay": ' || nvl(p_event.isAllDay,g_false) ||',' else '' end || 
                        case when p_event.isOnlineMeeting is not null then '"isOnlineMeeting": ' || nvl(p_event.isOnlineMeeting,g_false) ||',' else '' end ||
                        case when p_event.isOrganizer is not null then '"isOrganizer": ' || nvl(p_event.isOrganizer,g_true) ||',' else '' end ||
                        case when p_event.isReminderOn is not null then '"isReminderOn": ' || nvl(p_event.isReminderOn,g_true) ||',' else '' end ||
                        case when p_event.location is not null then '"location": {"displayName": "' || APEX_ESCAPE.JSON(p_event.location) || '"},' else '' end || 
                        case when p_event.organizer is not null then '"organizer": {"emailAddress": "' || APEX_ESCAPE.JSON(p_event.organizer) || '"},' else '' end || 
                        case when p_event.onlineMeetingProvider is not null then '"onlineMeetingProvider": ' || p_event.onlineMeetingProvider || ',' else '' end || 
                        case when p_event.onlineMeetingProvider is not null and p_event.onlineMeetingUrl is not null then '"onlineMeetingUrl": "' || p_event.onlineMeetingProvider || '",' else '' end || 
                        case when nvl(p_event.isReminderOn,g_true) = 'true' then '"reminderMinutesBeforeStart": ' || nvl(p_event.reminderMinutesBeforeStart,g_default_rmbs) || ',' else '' end || 
                        case when p_event.responseRequested is not null then '"responseRequested": ' || nvl(p_event.responseRequested,g_true) || ',' else '' end || 
                        case when p_event.sensitivity is not null then '"sensitivity": "' || nvl(p_event.responseRequested,g_sensitivity_normal) || '",' else '' end ||
                        case when p_event.showAs is not null then '"showAs": "' || nvl(p_event.showAs,g_showAs_free) || '",' else '' end ||
                        case when p_event.start_date is not null then '"start": {"dateTime": "' || get_date_str(p_event.start_date) || '","timeZone": "' || nvl(p_event.originalStartTimeZone,g_defaultTimeZone) || '"},' else '' end ||
                        case when p_event.event_type is not null then '"type": "' || nvl(p_event.event_type,g_event_type_singleInstance) || '",' else '' end ||
                        --Must re-specify the subject when event is updated
                        case when p_event.subject is not null then '"subject": "' || nvl(APEX_ESCAPE.JSON(p_event.subject),'No Subject') || '"' else '"subject": "' || 'No Subject given when updated"' end ||
                      '}';
    
    return l_returnvalue;

  end get_request_body;

begin
  
  IF p_user IS NOT NULL and p_calendar IS NOT NULL THEN
    l_url := g_url || 'users/' || p_user || '/calendars/' || p_calendar || '/events/' || p_event.event_id;
  ELSE 
    l_url := g_url || 'me/calendar/events/' || p_event.event_id;
  END IF;

  l_response := make_request(p_url => l_url,
                             p_http_method => l_http_method,
                             p_body => get_request_body);

  IF apex_web_service.g_status_code != 200 THEN
    --check_for_errors should have picked up everything but if not log it
    raise_error('Other Error: ' || apex_web_service.g_status_code || ': ' || l_response);
  END IF;
  
  return l_response;

end patch_cal_event;


procedure delete_cal_event(p_event_id in varchar2,
                           p_user in varchar2 := null,
                           p_calendar in varchar2 := null)
as

  l_http_method                  constant varchar2(30) := g_delete;
  l_url                          varchar2(2000);
  l_response                     CLOB;

begin
  
  IF p_user IS NOT NULL and p_calendar IS NOT NULL THEN
    l_url := g_url || 'users/' || p_user || '/calendars/' || p_calendar || '/events/' || p_event_id;
  ELSE 
    l_url := g_url || 'me/calendar/events/' || p_event_id;
  END IF;

  l_response := make_request(p_url => l_url,
                             p_http_method => l_http_method,
                             p_body => l_response);

  IF apex_web_service.g_status_code != 204 THEN
    --check_for_errors should have picked up everything but if not log it
    raise_error('Other Error: ' || apex_web_service.g_status_code || ': ' || l_response);
  END IF;

end delete_cal_event;

end ms_graph_util_pkg;
/

