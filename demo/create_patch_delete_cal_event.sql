declare 
l_event ms_graph_util_pkg.t_event;
l_token varchar2(2000);
begin

ms_graph_util_pkg.init(p_url => '', --e.g. https://graph.microsoft.com/v1.0/
                       p_login_url => '', --e.g. https://login.microsoftonline.com/
                       p_oauth2_url => '', --e.g. /oauth2/v2.0/token
                       p_login_scope => '', --e.g. https://graph.microsoft.com/.default
                       p_tenantid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientsecret => '', --e.g. abcdefghijklmnopqrstuvwxyz12345678
                       p_username => '', --e.g. API_USER@company.org
                       p_password => '', --e.g. password
                       p_wallet_path => '', --e.g. file:/u01/app/oracle/admin/dbname/wallet
                       p_wallet_password => ''); --e.g. password
 
 l_event.subject := ' Test Subject ' || to_char(sysdate);
 l_event.body := 'Test Body ' || to_char(sysdate);
 l_event.location := 'Location';
 l_event.start_date := sysdate+0.5;
 l_event.end_date := sysdate+0.51;
 l_event.isReminderOn := 'false';
 
 l_event.event_id := ms_graph_util_pkg.create_cal_event(p_event => l_event,
                                                        p_user => null,
                                                        p_calendar => null,
                                                        p_required_attendees => apex_util.string_to_table('abc@def.com:ghi@jkl.com'),
                                                        p_categories => apex_util.string_to_table('Misc.'));

 dbms_output.put_line(l_event.event_id);
 
end;

declare 
l_event ms_graph_util_pkg.t_event;
l_token varchar2(2000);
l_response CLOB;
begin

ms_graph_util_pkg.init(p_url => '', --e.g. https://graph.microsoft.com/v1.0/
                       p_login_url => '', --e.g. https://login.microsoftonline.com/
                       p_oauth2_url => '', --e.g. /oauth2/v2.0/token
                       p_login_scope => '', --e.g. https://graph.microsoft.com/.default
                       p_tenantid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientsecret => '', --e.g. abcdefghijklmnopqrstuvwxyz12345678
                       p_username => '', --e.g. API_USER@company.org
                       p_password => '', --e.g. password
                       p_wallet_path => '', --e.g. file:/u01/app/oracle/admin/dbname/wallet
                       p_wallet_password => ''); --e.g. password
 
 l_event.event_id := 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
 l_event.subject := 'Patched Test Subject ' || to_char(sysdate);
 
 l_response := ms_graph_util_pkg.patch_cal_event(p_event => l_event,
                                                 p_user => null,
                                                 p_calendar => null,
                                                 p_required_attendees => apex_util.string_to_table(''),
                                                 p_categories => apex_util.string_to_table(''));

 dbms_output.put_line(l_response);
 
end;


begin

ms_graph_util_pkg.init(p_url => '', --e.g. https://graph.microsoft.com/v1.0/
                       p_login_url => '', --e.g. https://login.microsoftonline.com/
                       p_oauth2_url => '', --e.g. /oauth2/v2.0/token
                       p_login_scope => '', --e.g. https://graph.microsoft.com/.default
                       p_tenantid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientid => '', --e.g. xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
                       p_clientsecret => '', --e.g. abcdefghijklmnopqrstuvwxyz12345678
                       p_username => '', --e.g. API_USER@company.org
                       p_password => '', --e.g. password
                       p_wallet_path => '', --e.g. file:/u01/app/oracle/admin/dbname/wallet
                       p_wallet_password => ''); --e.g. password

ms_graph_util_pkg.delete_cal_event('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',
                                   null,
                                   null);
end;