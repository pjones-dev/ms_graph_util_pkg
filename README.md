# ms_graph_util_pkg
Package handles Microsoft Graph REST API calls in PL/SQL which can be used to integrate with Office 365

At the moment package handles:

Create Event: https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0&tabs=http

Update Event: https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-1.0&tabs=http

Delete Event: https://docs.microsoft.com/en-us/graph/api/event-delete?view=graph-rest-1.0&tabs=http

Handles events in the 'me' calendar and also another user's i.e. shared calendar: populate p_user and p_calendar on the calls

Package was initially based on calendar calls in ms_ews_util_pkg package from alexandria-plsql-utils available at https://github.com/mortenbra/alexandria-plsql-utils. 

Updated to use REST calls (EWS is SOAP), OAUTH2 Token Authentication (ms_ews_util_pkg used basic auth) and events (known as items in EWS)

To do:
-Not everything in events is supported yet (e.g. Attachments)
-Add other functions available in Graph (e.g. mail, contacts)
