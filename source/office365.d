module office365;

import vibe.d;

///
struct Events
{
	///
	Event[] value;
}

///
struct EventDate
{
	///
	@name("DateTime")
	string _DateTime;
	///
	string TimeZone;
}

///
struct EventLocation
{
	///
	string DisplayName;
}

/++ 
 + Microsoft.OutlookServices.Event
 + see https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#EventResource
 +/
struct Event
{
	///
	string[] Categories;
	///
	bool IsReminderOn;
	///
	bool HasAttachments;
	///
	string Subject;
	///
	bool IsAllDay;
	///
	EventDate Start;
	///
	EventDate End;
	///
	EventLocation Location;
}

///
@path("/api/v2.0/me")
interface Office365Api
{
    /// uses https://outlook.office.com/api/v2.0/me/calendarview
    @method(HTTPMethod.GET)
    Events calendarview(DateTime startdatetime, DateTime enddatetime);
}

///
static Office365Api createOfficeApi(string access_token)
{
	import vibe.web.rest:RestInterfaceClient;
	import vibe.http.client:HTTPClientRequest;
	auto res = new RestInterfaceClient!Office365Api("https://outlook.office.com");

	res.requestFilter = (HTTPClientRequest req){
		req.headers["Authorization"] = "Bearer "~access_token;
	};

	return res;
}