import { runAppleScript } from "run-applescript";

const CONFIG = {
	MAX_REMINDERS: 50,
	MAX_LISTS: 20,
	TIMEOUT_MS: 8000,
};

const APPLESCRIPT_TIMEOUT_SECONDS = Math.max(1, Math.ceil(CONFIG.TIMEOUT_MS / 1000));

type AppleScriptRecord = Record<string, unknown>;

interface ReminderList {
	name: string;
	id: string;
}

interface Reminder {
	name: string;
	id: string;
	body: string;
	completed: boolean;
	dueDate: string | null;
	listName: string;
	completionDate?: string | null;
	creationDate?: string | null;
	modificationDate?: string | null;
	remindMeDate?: string | null;
	priority?: number;
}

interface OpenReminderResult {
	success: boolean;
	message: string;
	reminder?: Reminder;
}

const REMINDER_RECORD_HANDLER = `
on buildReminderRecord(reminderItem, listNameValue)
	set reminderName to ""
	try
		set reminderName to name of reminderItem as string
	end try

	set reminderBody to ""
	try
		set reminderBody to body of reminderItem as string
	on error
		set reminderBody to ""
	end try

	set reminderId to ""
	try
		set reminderId to id of reminderItem as string
	end try

	set reminderCompleted to false
	try
		set reminderCompleted to completed of reminderItem as boolean
	end try

	set dueString to ""
	try
		set dueValue to due date of reminderItem
		if dueValue is not missing value then
			set dueString to dueValue as string
		end if
	on error
		set dueString to ""
	end try

	set completionString to ""
	try
		set completionValue to completion date of reminderItem
		if completionValue is not missing value then
			set completionString to completionValue as string
		end if
	on error
		set completionString to ""
	end try

	set creationString to ""
	try
		set creationValue to creation date of reminderItem
		if creationValue is not missing value then
			set creationString to creationValue as string
		end if
	on error
		set creationString to ""
	end try

	set modificationString to ""
	try
		set modificationValue to modification date of reminderItem
		if modificationValue is not missing value then
			set modificationString to modificationValue as string
		end if
	on error
		set modificationString to ""
	end try

	set remindMeString to ""
	try
		set remindValue to remind me date of reminderItem
		if remindValue is not missing value then
			set remindMeString to remindValue as string
		end if
	on error
		set remindMeString to ""
	end try

	set priorityValue to ""
	try
		set priorityValue to priority of reminderItem as integer
	on error
		set priorityValue to ""
	end try

	return {name:reminderName, id:reminderId, body:reminderBody, completed:reminderCompleted, listName:listNameValue, dueDate:dueString, completionDate:completionString, creationDate:creationString, modificationDate:modificationString, remindMeDate:remindMeString, priority:priorityValue}
end buildReminderRecord
`;

function toAppleScriptString(value: string): string {
	return JSON.stringify(value ?? "");
}

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
	if (typeof value === "string") {
		const normalized = value.toLowerCase();
		if (normalized === "true" || normalized === "yes") {
			return true;
		}
		if (normalized === "false" || normalized === "no") {
			return false;
		}
	}
	if (typeof value === "number") {
		return value !== 0;
	}
	return fallback;
}

function coerceNumber(value: unknown): number | undefined {
	if (typeof value === "number" && Number.isFinite(value)) {
		return value;
	}
	if (typeof value === "string" && value.trim() !== "") {
		const parsed = Number(value);
		if (!Number.isNaN(parsed)) {
			return parsed;
		}
	}
	return undefined;
}

function toIsoString(value: unknown): string | null {
	if (!value) {
		return null;
	}
	if (value instanceof Date && !Number.isNaN(value.getTime())) {
		return value.toISOString();
	}
	const asString = coerceString(value);
	if (!asString) {
		return null;
	}
	const parsed = Date.parse(asString);
	if (Number.isNaN(parsed)) {
		return null;
	}
	return new Date(parsed).toISOString();
}

function mapReminderRecords(raw: unknown, fallbackListName?: string): Reminder[] {
	if (!Array.isArray(raw)) {
		const single = raw && typeof raw === "object" ? [raw] : [];
		return mapReminderRecords(single, fallbackListName);
	}

	return raw
		.filter((item): item is AppleScriptRecord => item !== null && typeof item === "object")
		.map((record) => {
			const name = coerceString(record.name, "Untitled Reminder");
			const id = coerceString(record.id, "");
			const body = coerceString(record.body, "");
			const completed = coerceBoolean(record.completed);
			const dueDate = toIsoString(record.dueDate);
			const listName = coerceString(record.listName, fallbackListName || "Reminders");
			const completionDate = toIsoString(record.completionDate);
			const creationDate = toIsoString(record.creationDate);
			const modificationDate = toIsoString(record.modificationDate);
			const remindMeDate = toIsoString(record.remindMeDate);
			const priority = coerceNumber(record.priority);

			return {
				name,
				id,
				body,
				completed,
				dueDate,
				listName,
				completionDate: completionDate ?? undefined,
				creationDate: creationDate ?? undefined,
				modificationDate: modificationDate ?? undefined,
				remindMeDate: remindMeDate ?? undefined,
				priority,
			};
		})
		.filter((reminder) => reminder.id || reminder.name);
}

async function ensureRemindersAccess(): Promise<void> {
	const accessResult = await requestRemindersAccess();
	if (!accessResult.hasAccess) {
		throw new Error(accessResult.message);
	}
}

type ReminderFetchOptions = {
	listName?: string;
	listId?: string;
	searchText?: string;
	maxReminders?: number;
};

async function fetchReminders({
	listName,
	listId,
	searchText,
	maxReminders = CONFIG.MAX_REMINDERS,
}: ReminderFetchOptions): Promise<Reminder[]> {
	await ensureRemindersAccess();

	const sanitizedListName = listName?.trim();
	const sanitizedListId = listId?.trim();
	const sanitizedSearch = searchText?.trim();
	const cappedLimit = Math.max(1, Math.min(maxReminders, CONFIG.MAX_REMINDERS));

	const hasSearch = Boolean(sanitizedSearch);
	const hasListId = Boolean(sanitizedListId);
	const hasListName = Boolean(sanitizedListName);

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Reminders"
		set reminderList to {}
		set reminderCount to 0
		set maxReminders to ${cappedLimit}
		set processedLists to 0
		set searchMode to ${hasSearch ? "true" : "false"}
		set searchLower to ""

		if searchMode then
			set searchNeedle to ${hasSearch ? toAppleScriptString(sanitizedSearch!) : "\"\""}
			set searchLower to do shell script "echo " & quoted form of searchNeedle & " | tr '[:upper:]' '[:lower:]'"
		end if

		set candidateLists to {}
		${hasListId
			? `set desiredId to ${toAppleScriptString(sanitizedListId!)}
		repeat with candidate in lists
			try
				if (id of candidate as string) is desiredId then
					set candidateLists to {candidate}
					exit repeat
				end if
			on error
				-- Ignore lookup errors
			end try
		end repeat`
			: hasListName
				? `set desiredName to ${toAppleScriptString(sanitizedListName!)}
		set desiredLower to do shell script "echo " & quoted form of desiredName & " | tr '[:upper:]' '[:lower:]'"
		repeat with candidate in lists
			try
				set candidateLower to do shell script "echo " & quoted form of (name of candidate as string) & " | tr '[:upper:]' '[:lower:]'"
				if candidateLower contains desiredLower or desiredLower contains candidateLower then
					set candidateLists to {candidate}
					exit repeat
				end if
			on error
				-- Ignore lookup errors
			end try
		end repeat`
				: "set candidateLists to lists"}

		if (count of candidateLists) is 0 then
			return {status:"error", reason:"list_not_found"}
		end if

		repeat with currentList in candidateLists
			if reminderCount >= maxReminders then exit repeat
			if processedLists >= ${CONFIG.MAX_LISTS} then exit repeat
			set processedLists to processedLists + 1
			try
				set listNameValue to name of currentList as string
				set reminderItems to reminders of currentList
				repeat with reminderItem in reminderItems
					if reminderCount >= maxReminders then exit repeat
					try
						set includeReminder to true
						if searchMode then
							set combinedText to ""
							try
								set combinedText to name of reminderItem as string
							end try
							try
								set combinedText to combinedText & " " & (body of reminderItem as string)
							on error
								-- Ignore body errors
							end try
							set combinedLower to do shell script "echo " & quoted form of combinedText & " | tr '[:upper:]' '[:lower:]'"
							if combinedLower does not contain searchLower then
								set includeReminder to false
							end if
						end if

						if includeReminder then
							set reminderRecord to my buildReminderRecord(reminderItem, listNameValue)
							set end of reminderList to reminderRecord
							set reminderCount to reminderCount + 1
						end if
					on error
						-- Ignore reminder level errors
					end try
				end repeat
			on error
				-- Ignore list level errors
			end try
		end repeat

		return reminderList
	end tell
end timeout
${REMINDER_RECORD_HANDLER}`;

	const rawResult = await runAppleScript(script);

	if (rawResult && typeof rawResult === "object" && !Array.isArray(rawResult)) {
		const status = coerceString((rawResult as AppleScriptRecord).status);
		if (status === "error") {
			const reason = coerceString((rawResult as AppleScriptRecord).reason);
			if (reason === "list_not_found") {
				throw new Error("Could not find the specified reminders list.");
			}
			throw new Error(reason || "An unknown Reminders error occurred.");
		}
	}

	return mapReminderRecords(rawResult, sanitizedListName);
}

async function showReminderById(reminderId: string): Promise<{ success: boolean; message: string }> {
	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Reminders"
		try
			set targetReminder to first reminder whose id is ${toAppleScriptString(reminderId)}
		on error errMsg
			return "ERROR:" & errMsg
		end try

		try
			show targetReminder
		on error
			try
				set parentList to container of targetReminder
				show parentList
			on error
				-- Best effort
			end try
		end try

		activate
		return "SUCCESS"
	end tell
end timeout`;

	const result = await runAppleScript(script);
	if (typeof result === "string" && result.startsWith("ERROR:")) {
		return { success: false, message: result.replace("ERROR:", "").trim() || "Failed to open reminder." };
	}
	return { success: true, message: "Reminders app opened." };
}

async function checkRemindersAccess(): Promise<boolean> {
	try {
		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Reminders"
		return name
	end tell
end timeout`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Reminders app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

async function requestRemindersAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		const hasAccess = await checkRemindersAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Reminders access is already granted.",
			};
		}

		return {
			hasAccess: false,
			message:
				"Reminders access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Automation\n2. Enable 'Reminders' for your terminal/app\n3. Restart your terminal and try again",
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Reminders access: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function getAllLists(): Promise<ReminderList[]> {
	try {
		await ensureRemindersAccess();

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Reminders"
		set listArray to {}
		set listCount to 0
		set allLists to lists

		repeat with currentList in allLists
			if listCount >= ${CONFIG.MAX_LISTS} then exit repeat
			try
				set listName to name of currentList as string
				set listId to id of currentList as string
				set end of listArray to {name:listName, id:listId}
				set listCount to listCount + 1
			on error
				-- Ignore lists we cannot read
			end try
		end repeat

		return listArray
	end tell
end timeout`;

		const result = await runAppleScript(script);
		const resultArray = Array.isArray(result) ? result : result ? [result] : [];

		return resultArray
			.filter((list): list is AppleScriptRecord => list !== null && typeof list === "object")
			.map((list) => ({
				name: coerceString(list.name, "Untitled List"),
				id: coerceString(list.id, ""),
			}))
			.filter((list) => Boolean(list.id || list.name));
	} catch (error) {
		console.error(
			`Error getting reminder lists: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getAllReminders(listName?: string): Promise<Reminder[]> {
	return fetchReminders({ listName });
}

async function searchReminders(searchText: string): Promise<Reminder[]> {
	if (!searchText || searchText.trim() === "") {
		return [];
	}

	return fetchReminders({ searchText });
}

async function createReminder(
	name: string,
	listName: string = "Reminders",
	notes?: string,
	dueDate?: string,
): Promise<Reminder> {
	await ensureRemindersAccess();

	if (!name || name.trim() === "") {
		throw new Error("Reminder name cannot be empty");
	}

	const trimmedName = name.trim();
	const trimmedNotes = notes?.trim();
	const trimmedList = listName?.trim();

	let dueTimestamp: number | null = null;
	if (dueDate) {
		const parsedDueDate = new Date(dueDate);
		if (Number.isNaN(parsedDueDate.getTime())) {
			throw new Error("Invalid due date format. Please use ISO date strings.");
		}
		dueTimestamp = Math.round(parsedDueDate.getTime() / 1000);
	}

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Reminders"
		set reminderName to ${toAppleScriptString(trimmedName)}
		set targetList to missing value

		${trimmedList
			? `set desiredName to ${toAppleScriptString(trimmedList)}
		try
			set targetList to first list whose name is desiredName
		on error
			repeat with candidate in lists
				try
					if (id of candidate as string) is desiredName then
						set targetList to candidate
						exit repeat
					end if
				on error
					-- Ignore lookup errors
				end try
			end repeat
		end try`
			: ""}

		if targetList is missing value then
			try
				set targetList to default list
			on error
				set targetList to first list
			end try
		end if

		set newReminder to make new reminder at targetList with properties {name:reminderName}

		${trimmedNotes ? `set body of newReminder to ${toAppleScriptString(trimmedNotes)}` : ""}

		${dueTimestamp !== null
			? `try
				set desiredTimestamp to ${dueTimestamp}
				set nowUnix to (do shell script "date +%s") as integer
				set offsetSeconds to desiredTimestamp - nowUnix
				set dueDateValue to (current date) + offsetSeconds
				set due date of newReminder to dueDateValue
			on error
				-- Ignore due date failures
			end try`
			: ""}

		return my buildReminderRecord(newReminder, name of targetList as string)
	end tell
end timeout
${REMINDER_RECORD_HANDLER}`;

	const rawResult = await runAppleScript(script);
	const reminders = mapReminderRecords(rawResult, trimmedList);
	if (reminders.length === 0) {
		throw new Error("Failed to create reminder.");
	}

	return reminders[0];
}

async function openReminder(searchText: string): Promise<OpenReminderResult> {
	try {
		if (!searchText || searchText.trim() === "") {
			return { success: false, message: "Search text cannot be empty." };
		}

		const reminders = await fetchReminders({ searchText, maxReminders: CONFIG.MAX_REMINDERS });
		if (reminders.length === 0) {
			return { success: false, message: "No matching reminders found." };
		}

		const reminder = reminders[0];
		const openResult = await showReminderById(reminder.id);

		if (!openResult.success) {
			return { success: false, message: openResult.message };
		}

		return {
			success: true,
			message: openResult.message,
			reminder,
		};
	} catch (error) {
		return {
			success: false,
			message: `Failed to open reminder: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function getRemindersFromListById(
	listId: string,
	_props?: string[],
): Promise<Reminder[]> {
	if (!listId || listId.trim() === "") {
		return [];
	}

	return fetchReminders({ listId });
}

export default {
	getAllLists,
	getAllReminders,
	searchReminders,
	createReminder,
	openReminder,
	getRemindersFromListById,
	requestRemindersAccess,
};
