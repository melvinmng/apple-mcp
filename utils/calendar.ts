import { runAppleScript } from "run-applescript";

interface CalendarEvent {
	id: string;
	title: string;
	location: string | null;
	notes: string | null;
	startDate: string | null;
	endDate: string | null;
	calendarName: string;
	isAllDay: boolean;
	url: string | null;
}

const CONFIG = {
	TIMEOUT_MS: 10000,
	MAX_EVENTS: 50,
};

const APPLESCRIPT_TIMEOUT_SECONDS = Math.max(1, Math.ceil(CONFIG.TIMEOUT_MS / 1000));

const ONE_DAY_MS = 24 * 60 * 60 * 1000;

const EVENT_SCRIPT_HELPERS = `
on parseIsoDate(isoString)
	try
		set command to "date -jf '%Y-%m-%dT%H:%M:%S%z' " & quoted form of isoString & " '+%Y-%m-%d %H:%M:%S'"
		set formatted to do shell script command
		return date formatted
	on error
		return current date
	end try
end parseIsoDate

on formatDateToUnix(inputDate)
	try
		set nowDate to current date
		set nowUnix to (do shell script "date +%s") as real
		set deltaSeconds to inputDate - nowDate
		set unixValue to nowUnix + deltaSeconds
		return unixValue as string
	on error
		return ""
	end try
end formatDateToUnix

on buildEventRecord(eventItem, calendarNameValue)
	set eventSummary to ""
	try
		set eventSummary to summary of eventItem as string
	end try

	set eventLocation to ""
	try
		set eventLocation to location of eventItem as string
	end try

	set eventNotes to ""
	try
		set eventNotes to description of eventItem as string
	end try

	set eventUrl to ""
	try
		set eventUrl to url of eventItem as string
	end try

	set startValue to ""
	try
		set startValue to my formatDateToUnix(start date of eventItem)
	end try

	set endValue to ""
	try
		set endValue to my formatDateToUnix(end date of eventItem)
	end try

	set eventUid to ""
	try
		set eventUid to uid of eventItem as string
	end try

	set isAllDayValue to false
	try
		set isAllDayValue to allday event of eventItem as boolean
	end try

	return {id:eventUid, title:eventSummary, location:eventLocation, notes:eventNotes, startDate:startValue, endDate:endValue, calendarName:calendarNameValue, isAllDay:isAllDayValue, url:eventUrl}
end buildEventRecord
`;

type AppleScriptRecord = Record<string, unknown>;

function coerceString(value: unknown, fallback = ""): string {
	if (typeof value === "string") {
		return value;
	}
	if (value === null || value === undefined) {
		return fallback;
	}
	return String(value);
}

function coerceBoolean(value: unknown, fallback = false): boolean {
	if (typeof value === "boolean") {
		return value;
	}
	if (typeof value === "number") {
		return value !== 0;
	}
	if (typeof value === "string") {
		const normalized = value.toLowerCase();
		if (["true", "yes", "1"].includes(normalized)) {
			return true;
		}
		if (["false", "no", "0"].includes(normalized)) {
			return false;
		}
	}
	return fallback;
}

function toNumber(value: unknown): number | null {
	if (typeof value === "number" && Number.isFinite(value)) {
		return value;
	}
	if (typeof value === "string" && value.trim() !== "") {
		const parsed = Number(value);
		if (!Number.isNaN(parsed)) {
			return parsed;
		}
	}
	return null;
}

function toIsoFromUnix(value: unknown): string | null {
	const unixSeconds = toNumber(value);
	if (unixSeconds === null) {
		return null;
	}
	return new Date(unixSeconds * 1000).toISOString();
}

function parseDateInput(value?: string): Date | null {
	if (!value) {
		return null;
	}
	const parsed = new Date(value);
	if (Number.isNaN(parsed.getTime())) {
		return null;
	}
	return parsed;
}

function ensureDateRange(
	start: Date | null,
	end: Date | null,
	fallbackDays: number,
): { start: Date; end: Date } {
	const startDate = start ? new Date(start) : new Date();
	const defaultEnd = new Date(startDate.getTime() + fallbackDays * ONE_DAY_MS);
	const endCandidate = end ? new Date(end) : defaultEnd;

	if (endCandidate.getTime() <= startDate.getTime()) {
		return {
			start: startDate,
			end: new Date(startDate.getTime() + Math.max(1, fallbackDays) * ONE_DAY_MS),
		};
	}

	return { start: startDate, end: endCandidate };
}

function toAppleScriptIso(date: Date): string {
	const year = date.getFullYear();
	const month = String(date.getMonth() + 1).padStart(2, "0");
	const day = String(date.getDate()).padStart(2, "0");
	const hours = String(date.getHours()).padStart(2, "0");
	const minutes = String(date.getMinutes()).padStart(2, "0");
	const seconds = String(date.getSeconds()).padStart(2, "0");
	const offsetMinutes = date.getTimezoneOffset();
	const sign = offsetMinutes > 0 ? "-" : "+";
	const absMinutes = Math.abs(offsetMinutes);
	const offsetHours = String(Math.floor(absMinutes / 60)).padStart(2, "0");
	const offsetMins = String(absMinutes % 60).padStart(2, "0");

	return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${sign}${offsetHours}${offsetMins}`;
}

function getEventSortKey(event: CalendarEvent): number {
	if (event.startDate) {
		const parsed = Date.parse(event.startDate);
		if (!Number.isNaN(parsed)) {
			return parsed;
		}
	}
	if (event.endDate) {
		const parsed = Date.parse(event.endDate);
		if (!Number.isNaN(parsed)) {
			return parsed;
		}
	}
	return Number.MAX_SAFE_INTEGER;
}

function mapEventRecords(raw: unknown): CalendarEvent[] {
	const records: AppleScriptRecord[] = Array.isArray(raw)
		? raw.filter((item): item is AppleScriptRecord => item !== null && typeof item === "object")
		: [];

	const deduped = new Map<string, CalendarEvent>();
	const fallback: CalendarEvent[] = [];

	for (const record of records) {
		const id = coerceString(record.id);
		const title = coerceString(record.title, "Untitled Event");
		const location = coerceString(record.location);
		const notes = coerceString(record.notes);
		const calendarName = coerceString(record.calendarName, "Unknown Calendar");
		const url = coerceString(record.url);
		const startDate = toIsoFromUnix(record.startDate);
		const endDate = toIsoFromUnix(record.endDate);
		const event: CalendarEvent = {
			id: id || `${title}-${startDate ?? "unknown"}-${calendarName}`,
			title,
			location: location || null,
			notes: notes || null,
			startDate,
			endDate,
			calendarName,
			isAllDay: coerceBoolean(record.isAllDay),
			url: url || null,
		};

		if (id) {
			deduped.set(id, event);
		} else {
			fallback.push(event);
		}
	}

	const ordered = [...deduped.values(), ...fallback];
	ordered.sort((a, b) => getEventSortKey(a) - getEventSortKey(b));
	return ordered;
}

async function checkCalendarAccess(): Promise<boolean> {
	try {
		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Calendar"
		return name
	end tell
end timeout`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Calendar app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

async function requestCalendarAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		const hasAccess = await checkCalendarAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Calendar access is already granted.",
			};
		}

		return {
			hasAccess: false,
			message:
				"Calendar access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Calendars\n2. Enable access for your terminal/application\n3. Also check System Settings > Privacy & Security > Automation and allow 'Calendar'\n4. Restart your terminal and try again",
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Calendar access: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function ensureCalendarAccess(): Promise<void> {
	const access = await requestCalendarAccess();
	if (!access.hasAccess) {
		throw new Error(access.message);
	}
}

interface EventFetchOptions {
	fromDate: Date;
	toDate: Date;
	limit: number;
	calendarName?: string;
	searchText?: string;
}

async function fetchEvents({
	fromDate,
	toDate,
	limit,
	calendarName,
	searchText,
}: EventFetchOptions): Promise<CalendarEvent[]> {
	await ensureCalendarAccess();

	const maxEvents = Math.max(1, Math.min(limit, CONFIG.MAX_EVENTS));
	const sanitizedCalendar = calendarName?.trim();
	const sanitizedSearch = searchText?.trim();

	const startIso = toAppleScriptIso(fromDate);
	const endIso = toAppleScriptIso(toDate);

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	using terms from application "Calendar"
	tell application "Calendar"
		set maxEvents to ${maxEvents}
		set eventList to {}
		set eventCount to 0

		set startIso to ${JSON.stringify(startIso)}
		set endIso to ${JSON.stringify(endIso)}
		set startDateValue to my parseIsoDate(startIso)
		set endDateValue to my parseIsoDate(endIso)

		set searchMode to ${sanitizedSearch ? "true" : "false"}
		set searchLower to ${sanitizedSearch ? JSON.stringify(sanitizedSearch.toLowerCase()) : "\"\""}

		set calendarMode to ${sanitizedCalendar ? "true" : "false"}
		set calendarLower to ${sanitizedCalendar ? JSON.stringify(sanitizedCalendar.toLowerCase()) : "\"\""}

		set candidateCalendars to {}
		if calendarMode then
			repeat with calItem in calendars
				try
					set calNameValue to name of calItem as string
					set calLower to do shell script "echo " & quoted form of calNameValue & " | tr '[:upper:]' '[:lower:]'"
					if calLower contains calendarLower or calendarLower contains calLower then
						set end of candidateCalendars to calItem
					end if
				on error
					-- Ignore calendar lookup errors
				end try
			end repeat
		else
			set candidateCalendars to calendars
		end if

		if (count of candidateCalendars) is 0 then
			return {status:"error", reason:"calendar_not_found"}
		end if

		repeat with currentCalendar in candidateCalendars
			if eventCount >= maxEvents then exit repeat
			try
				set calNameValue to name of currentCalendar as string
				set eventItems to every event of currentCalendar whose start date ≤ endDateValue and end date ≥ startDateValue

				repeat with eventItem in eventItems
					if eventCount >= maxEvents then exit repeat
					try
						set includeEvent to true
						if searchMode then
							set combinedText to ""
							try
								set combinedText to summary of eventItem as string
							end try
							try
								set combinedText to combinedText & " " & location of eventItem as string
							end try
							try
								set combinedText to combinedText & " " & description of eventItem as string
							end try
							set combinedLower to do shell script "echo " & quoted form of combinedText & " | tr '[:upper:]' '[:lower:]'"
							if combinedLower does not contain searchLower then
								set includeEvent to false
							end if
						end if

						if includeEvent then
							set recordValue to my buildEventRecord(eventItem, calNameValue)
							set end of eventList to recordValue
							set eventCount to eventCount + 1
						end if
					on error
						-- Skip problematic events
					end try
				end repeat
			on error
				-- Skip calendars that cannot be read
			end try
		end repeat

		return eventList
	end tell
end timeout
${EVENT_SCRIPT_HELPERS}`;

	const rawResult = await runAppleScript(script);

	if (rawResult && typeof rawResult === "object" && !Array.isArray(rawResult)) {
		const status = coerceString((rawResult as AppleScriptRecord).status);
		if (status === "error") {
			const reason = coerceString((rawResult as AppleScriptRecord).reason);
			if (reason === "calendar_not_found") {
				throw new Error("Could not find a calendar matching the requested name.");
			}
			throw new Error(reason || "Calendar operation failed.");
		}
	}

	const events = mapEventRecords(rawResult);
	const startMs = fromDate.getTime();
	const endMs = toDate.getTime();

	const filtered = events.filter((event) => {
		const start = event.startDate ? Date.parse(event.startDate) : Number.NaN;
		const end = event.endDate ? Date.parse(event.endDate) : Number.NaN;
		const effectiveStart = Number.isNaN(start) ? end : start;
		const effectiveEnd = Number.isNaN(end) ? effectiveStart : end;

		if (Number.isNaN(effectiveStart) && Number.isNaN(effectiveEnd)) {
			return true;
		}

		return effectiveEnd >= startMs && effectiveStart <= endMs;
	});

	return filtered.slice(0, maxEvents);
}

async function getEvents(
	limit = 10,
	fromDate?: string,
	toDate?: string,
): Promise<CalendarEvent[]> {
	const parsedFrom = parseDateInput(fromDate);
	const parsedTo = parseDateInput(toDate);
	const { start, end } = ensureDateRange(parsedFrom, parsedTo, 7);

	return fetchEvents({
		fromDate: start,
		toDate: end,
		limit: Math.max(1, limit ?? 10),
	});
}

async function searchEvents(
	searchText: string,
	limit = 10,
	fromDate?: string,
	toDate?: string,
): Promise<CalendarEvent[]> {
	const trimmed = searchText.trim();
	if (!trimmed) {
		return [];
	}

	const parsedFrom = parseDateInput(fromDate);
	const parsedTo = parseDateInput(toDate);
	const { start, end } = ensureDateRange(parsedFrom, parsedTo, 30);

	return fetchEvents({
		fromDate: start,
		toDate: end,
		limit: Math.max(1, limit ?? 10),
		searchText: trimmed,
	});
}

async function createEvent(
	title: string,
	startDate: string,
	endDate: string,
	location?: string,
	notes?: string,
	isAllDay = false,
	calendarName?: string,
): Promise<{ success: boolean; message: string; eventId?: string; event?: CalendarEvent }> {
	await ensureCalendarAccess();

	const trimmedTitle = title?.trim();
	if (!trimmedTitle) {
		return { success: false, message: "Event title cannot be empty." };
	}

	const start = parseDateInput(startDate);
	const end = parseDateInput(endDate);
	if (!start || !end) {
		return { success: false, message: "Invalid start or end date. Please provide ISO-formatted dates." };
	}

	let startValue = new Date(start);
	let endValue = new Date(end);

	if (isAllDay) {
		const allDayStart = new Date(startValue);
		allDayStart.setHours(0, 0, 0, 0);
		const allDayEnd = new Date(endValue);
		allDayEnd.setHours(0, 0, 0, 0);
		if (allDayEnd.getTime() <= allDayStart.getTime()) {
			allDayEnd.setTime(allDayStart.getTime() + ONE_DAY_MS);
		}
		startValue = allDayStart;
		endValue = allDayEnd;
	}

	if (endValue.getTime() <= startValue.getTime()) {
		return { success: false, message: "End date must be after start date." };
	}

	const startIso = toAppleScriptIso(startValue);
	const endIso = toAppleScriptIso(endValue);
	const sanitizedCalendar = calendarName?.trim();
	const sanitizedLocation = location?.trim();
	const sanitizedNotes = notes?.trim();

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	using terms from application "Calendar"
	tell application "Calendar"
		set titleValue to ${JSON.stringify(trimmedTitle)}
		set startIso to ${JSON.stringify(startIso)}
		set endIso to ${JSON.stringify(endIso)}
		set isAllDayValue to ${isAllDay ? "true" : "false"}
		set locationValue to ${sanitizedLocation ? JSON.stringify(sanitizedLocation) : "\"\""}
		set notesValue to ${sanitizedNotes ? JSON.stringify(sanitizedNotes) : "\"\""}
		set calendarMode to ${sanitizedCalendar ? "true" : "false"}
		set calendarLower to ${sanitizedCalendar ? JSON.stringify(sanitizedCalendar.toLowerCase()) : "\"\""}

		set startDateValue to my parseIsoDate(startIso)
		set endDateValue to my parseIsoDate(endIso)

		set targetCalendar to missing value
		if calendarMode then
			repeat with calItem in calendars
				try
					set calNameValue to name of calItem as string
					set calLower to do shell script "echo " & quoted form of calNameValue & " | tr '[:upper:]' '[:lower:]'"
					if calLower contains calendarLower or calendarLower contains calLower then
						set targetCalendar to calItem
						exit repeat
					end if
				on error
					-- Ignore calendar lookup errors
				end try
			end repeat
		end if

		if targetCalendar is missing value then
			try
				set targetCalendar to first calendar
			on error errMsg
				return {status:"error", reason:errMsg}
			end try
		end if

		set newEvent to make new event at targetCalendar with properties {summary:titleValue, start date:startDateValue, end date:endDateValue, allday event:isAllDayValue}

		if locationValue is not "" then
			try
				set location of newEvent to locationValue
			on error
				-- Ignore location errors
			end try
		end if

		if notesValue is not "" then
			try
				set description of newEvent to notesValue
			on error
				-- Ignore notes errors
			end try
		end if

		return {status:"success", eventRecord:my buildEventRecord(newEvent, name of targetCalendar as string)}
	end tell
end timeout
${EVENT_SCRIPT_HELPERS}`;

	const rawResult = await runAppleScript(script);

	if (!rawResult || typeof rawResult !== "object" || Array.isArray(rawResult)) {
		return { success: false, message: "Failed to create event: no response from Calendar." };
	}

	const status = coerceString((rawResult as AppleScriptRecord).status);
	if (status !== "success") {
		const reason = coerceString((rawResult as AppleScriptRecord).reason);
		return { success: false, message: reason || "Failed to create event." };
	}

	const eventRecord = (rawResult as AppleScriptRecord).eventRecord;
	const mapped = mapEventRecords(eventRecord);
	const created = mapped[0];

	return {
		success: true,
		message: `Event "${trimmedTitle}" created successfully.`,
		eventId: created?.id,
		event: created,
	};
}

async function openEvent(eventId: string): Promise<{ success: boolean; message: string }> {
	await ensureCalendarAccess();
	const trimmedId = eventId?.trim();
	if (!trimmedId) {
		return { success: false, message: "Event ID is required." };
	}

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	using terms from application "Calendar"
	tell application "Calendar"
		try
			set targetEvent to first event whose uid is ${JSON.stringify(trimmedId)}
		on error
			return {status:"error", reason:"event_not_found"}
		end try

		try
			show targetEvent
			activate
			return {status:"success"}
		on error errMsg
			activate
			return {status:"error", reason:errMsg}
		end try
	end tell
end timeout`;

	const rawResult = await runAppleScript(script);

	if (!rawResult || typeof rawResult !== "object" || Array.isArray(rawResult)) {
		return { success: false, message: "Unable to open event." };
	}

	const status = coerceString((rawResult as AppleScriptRecord).status);
	if (status === "success") {
		return { success: true, message: "Calendar app opened to the requested event." };
	}

	const reason = coerceString((rawResult as AppleScriptRecord).reason);
	if (reason === "event_not_found") {
		return { success: false, message: "Event not found." };
	}

	return { success: false, message: reason || "Unable to open event." };
}

const calendar = {
	searchEvents,
	openEvent,
	getEvents,
	createEvent,
	requestCalendarAccess,
};

export default calendar;
