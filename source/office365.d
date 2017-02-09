module office365;

import vibe.d;

///
struct Events
{
	///
	Event[] value;
}

///
enum DefaultFolder
{
	Inbox = "inbox",
	SentItems = "sentitems",
	Drafts = "drafts",
}

///
struct Messages
{
	///
	Message[] value;
	///
	@optional
	@name("@odata.count")
	int count;
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
	bool IsCancelled;
	///
	bool IsOrganizer;
	///
	@byName EventImportance Importance;
	///
	EventDate Start;
	///
	EventDate End;
	///
	EventLocation Location;
	///
	EventResponseStatus ResponseStatus;
	///
	EventAttendees[] Attendees;
	///
	@optional
	string OnlineMeetingUrl;

}

///
struct EventAttendees
{
	///
	EventAttendeesType Type;
	///
	EventResponseStatus Status;
	///
	@name("EmailAddress") 
	EmailAddress _EmailAddress;
}

///
enum EventAttendeesType
{
	Required,
	Optional
}

///
enum EventImportance
{
	Normal,
	High,
	Low
}

///
enum ResponseState
{
	Organizer,
	None,
	NotResponded,
	Accepted,
	Declined
}

///
struct EventResponseStatus 
{
	///
	@byName ResponseState Response;
	///
	string Time;
}


///
struct MessageBody
{
	///
	string ContentType;
	///
	string Content;
}

///
enum Select
{
	Subject,
	BodyPreview,
	SentDateTime,
	Body,
	Sender,
	From,
}

///
struct SelectBuilder
{
	private string content;

	///
	auto add(Select selector)
	{
		import std.conv:to;

		if(content.length!=0)
			content ~= ",";

		content ~= to!string(selector);

		return this;
	}

	///
	string toString() const
	{
		return content;
	}
}

///
struct EmailAddress
{
	///
	struct NameAdressTuple
	{
		///
		string Name;
		///
		string Address;
	}

	///
	NameAdressTuple EmailAddress;
}

///
struct Message
{
	///
	string Id;
	///
	@optional
	string Subject;
	///
	@optional
	MessageBody Body;
	///
	@optional
	string BodyPreview;
	///
	@optional
	string SentDateTime;
	///
	@optional
	EmailAddress Sender;
	///
	@optional
	EmailAddress From;
	///
	@optional
	bool IsRead;
}

///
@path("/api/v2.0/me")
interface Office365Api
{
	/// uses https://outlook.office.com/api/v2.0/me/calendarview
	@method(HTTPMethod.GET)
	Events calendarview(DateTime startdatetime, DateTime enddatetime);

	@queryParam("_filter", "$filter")
	@queryParam("_top", "$top")
	@queryParam("_select", "$select")
	@path("/mailfolders/:folder/messages")
	@method(HTTPMethod.GET)
	{
		@queryParam("_count", "$count")
		Messages messages(string _folder, bool _count, string _filter, string _select="", int _top=10);
		Messages messages(string _folder, string _filter="", string _select="", int _top=10);
	}

	@path("/messages/:id")
	@method(HTTPMethod.DELETE)
	void deleteMessage(string _id);

	@path("/messages/:id")
	@method(HTTPMethod.PATCH)
	{
		Message updateRead(string _id, bool IsRead);
		Message updateCategories(string _id, string[] Categories);
	}
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