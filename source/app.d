import std.stdio;
import std.datetime;
import std.process;

import vibe.d;

import office365;

void main()
{
	// grab your Access token on https://oauthplay.azurewebsites.net
	auto token = environment["OFFICE365_TOKEN"];

	auto api = createOfficeApi(token);

	immutable time = Clock.currTime;
	DateTime start = cast(DateTime)time;
	start.timeOfDay = TimeOfDay.init;
	auto to = start + days(1);
	
	auto res = api.calendarview(start, to);

	writefln("events: %s\n",res);

	Messages resMail = api.messages(DefaultFolder.Inbox, "isread eq false", 
		SelectBuilder()
		.add(Select.From)
		.add(Select.Subject).toString);

	writefln("inbox: %s\n", serializeToJson(resMail).toPrettyString);

	assert(!resMail.value[0].IsRead);
	auto updateRes = api.updateRead(resMail.value[0].Id,true);
	assert(updateRes.IsRead);
	updateRes = api.updateRead(resMail.value[0].Id,false);
	assert(!updateRes.IsRead);
	writefln("update: %s\n", serializeToJson(updateRes).toPrettyString);
}