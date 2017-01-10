import std.stdio;
import std.datetime;
import std.process;

import vibe.d;

import office365;

void main()
{
	// grab your auth token on https://oauthplay.azurewebsites.net
	auto token = environment["OFFICE365_TOKEN"];

	auto api = createOfficeApi(token);

	auto time = Clock.currTime;
	DateTime start = cast(DateTime)time;
	start.timeOfDay = TimeOfDay.init;
	auto to = start + days(1);
	
	auto res = api.calendarview(start, to);

	writefln("res: %s",res);
}