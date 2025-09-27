import { runAppleScript } from "run-applescript";
import { writeFile, unlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";

const CONFIG = {
	MAX_NOTES: 50,
	MAX_CONTENT_PREVIEW: 200,
	TIMEOUT_MS: 8000,
};

const APPLESCRIPT_TIMEOUT_SECONDS = Math.max(1, Math.ceil(CONFIG.TIMEOUT_MS / 1000));

const NOTE_SCRIPT_HELPERS = `
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

on buildNoteRecord(noteItem, folderNameValue, previewLimit)
	set noteTitle to ""
	try
		set noteTitle to name of noteItem as string
	end try

	set noteContent to ""
	try
		set noteContent to plaintext of noteItem as string
		if previewLimit > 0 and (length of noteContent) > previewLimit then
			set noteContent to (text 1 thru previewLimit of noteContent) & "..."
		end if
	end try

	set creationValue to ""
	try
		set creationValue to my formatDateToUnix(creation date of noteItem)
	end try

	set modificationValue to ""
	try
		set modificationValue to my formatDateToUnix(modification date of noteItem)
	end try

	set noteId to ""
	try
		set noteId to id of noteItem as string
	end try

	return {name:noteTitle, content:noteContent, folderName:folderNameValue, creationDate:creationValue, modificationDate:modificationValue, id:noteId}
end buildNoteRecord
`;

type AppleScriptPrimitive = string | number | boolean | null;
type AppleScriptValue =
	| AppleScriptPrimitive
	| AppleScriptValue[]
	| { [key: string]: AppleScriptValue };

class AppleScriptSourceParser {
	private index = 0;

	constructor(private readonly input: string) {}

	parse(): AppleScriptValue {
		this.skipWhitespace();
		const value = this.parseValue();
		this.skipWhitespace();
		if (this.index < this.input.length) {
			throw new Error("Unexpected trailing content in AppleScript result");
		}
		return value;
	}

	private parseValue(): AppleScriptValue {
		this.skipWhitespace();
		const char = this.peek();

		if (!char) {
			throw new Error("Unexpected end of AppleScript result");
		}

		if (char === "{") {
			return this.parseCollection();
		}
		if (char === "\"") {
			return this.parseString();
		}
		if (this.startsWith("missing value")) {
			this.index += "missing value".length;
			return null;
		}
		if (this.startsWith("true")) {
			this.index += "true".length;
			return true;
		}
		if (this.startsWith("false")) {
			this.index += "false".length;
			return false;
		}

		if (char === "-" || this.isDigit(char)) {
			return this.parseNumber();
		}

		return this.parseBareword();
	}

	private parseCollection(): AppleScriptValue {
		this.consume("{");
		this.skipWhitespace();
		if (this.peek() === "}") {
			this.consume("}");
			return [];
		}

		const isRecord = this.isRecordCollection();

		if (isRecord) {
			const result: Record<string, AppleScriptValue> = {};
			while (true) {
				const key = this.parseKey();
				this.skipWhitespace();
				this.consume(":");
				const value = this.parseValue();
				result[key] = value;
				this.skipWhitespace();
				const next = this.peek();
				if (next === ",") {
					this.index += 1;
					continue;
				}
				break;
			}
			this.skipWhitespace();
			this.consume("}");
			return result;
		}

		const list: AppleScriptValue[] = [];
		while (true) {
			const value = this.parseValue();
			list.push(value);
			this.skipWhitespace();
			const next = this.peek();
			if (next === ",") {
				this.index += 1;
				continue;
			}
			break;
		}
		this.skipWhitespace();
		this.consume("}");
		return list;
	}

	private isRecordCollection(): boolean {
		let depth = 0;
		let position = this.index;
		while (position < this.input.length) {
			const char = this.input[position];
			if (this.isWhitespaceChar(char)) {
				position += 1;
				continue;
			}
			if (char === "\"") {
				position = this.skipStringFrom(position);
				continue;
			}
			if (char === "{") {
				depth += 1;
				position += 1;
				continue;
			}
			if (char === "}") {
				if (depth === 0) {
					return false;
				}
				depth -= 1;
				position += 1;
				continue;
			}
			if (depth === 0 && char === ":") {
				return true;
			}
			if (depth === 0 && char === ",") {
				return false;
			}
			position += 1;
		}
		return false;
	}

	private parseKey(): string {
		this.skipWhitespace();
		const char = this.peek();
		if (!char) {
			throw new Error("Unexpected end while parsing key");
		}
		if (char === "\"") {
			return this.parseString();
		}
		const start = this.index;
		while (true) {
			const current = this.peek();
			if (!current || !/[A-Za-z0-9_]/.test(current)) {
				break;
			}
			this.index += 1;
		}
		if (start === this.index) {
			throw new Error("Invalid record key in AppleScript result");
		}
		return this.input.slice(start, this.index);
	}

	private parseString(): string {
		this.consume("\"");
		let result = "";
		while (this.index < this.input.length) {
			const char = this.consume();
			if (char === "\"") {
				return result;
			}
			if (char === "\\") {
				const next = this.consume();
				switch (next) {
					case "\"":
						result += "\"";
						break;
					case "\\":
						result += "\\";
						break;
					case "n":
						result += "\n";
						break;
					case "r":
						result += "\r";
						break;
					case "t":
						result += "\t";
						break;
					default:
						result += next;
				}
				continue;
			}
			result += char;
		}
		throw new Error("Unterminated string literal in AppleScript result");
	}

	private parseNumber(): number {
		const start = this.index;
		if (this.peek() === "-") {
			this.index += 1;
		}
		while (this.isDigit(this.peek())) {
			this.index += 1;
		}
		if (this.peek() === ".") {
			this.index += 1;
			while (this.isDigit(this.peek())) {
				this.index += 1;
			}
		}
		const text = this.input.slice(start, this.index);
		return Number(text);
	}

	private parseBareword(): string {
		const start = this.index;
		while (true) {
			const current = this.peek();
			if (!current || /[,}]/.test(current) || this.isWhitespaceChar(current)) {
				break;
			}
			this.index += 1;
		}
		if (start === this.index) {
			throw new Error("Unable to parse AppleScript value");
		}
		return this.input.slice(start, this.index);
	}

	private skipWhitespace(): void {
		while (this.isWhitespaceChar(this.peek())) {
			this.index += 1;
		}
	}

	private skipStringFrom(position: number): number {
		let index = position + 1;
		while (index < this.input.length) {
			const char = this.input[index];
			if (char === "\\") {
				index += 2;
				continue;
			}
			if (char === "\"") {
				return index + 1;
			}
			index += 1;
		}
		throw new Error("Unterminated string literal detected while scanning");
	}

	private consume(expected?: string): string {
		if (this.index >= this.input.length) {
			throw new Error("Unexpected end of AppleScript result");
		}
		const char = this.input[this.index];
		if (expected && char !== expected) {
			throw new Error(`Expected "${expected}" but found "${char}" in AppleScript result`);
		}
		this.index += 1;
		return char;
	}

	private peek(): string | undefined {
		return this.input[this.index];
	}

	private startsWith(value: string): boolean {
		return this.input.slice(this.index, this.index + value.length) === value;
	}

	private isDigit(char: string | undefined): boolean {
		return !!char && /[0-9]/.test(char);
	}

	private isWhitespaceChar(char: string | undefined): boolean {
		return !!char && /\s/.test(char);
	}
}

function parseAppleScriptResult(raw: unknown): AppleScriptValue {
	if (raw === null || raw === undefined) {
		return null;
	}
	if (Array.isArray(raw)) {
		return raw as AppleScriptValue;
	}
	if (typeof raw === "object") {
		return raw as AppleScriptValue;
	}
	if (typeof raw !== "string") {
		throw new Error("Unsupported AppleScript return type");
	}

	const trimmed = raw.trim();
	if (!trimmed) {
		return null;
	}

	const normalized = trimmed.replace(/^=>\s*/, "");
	const parser = new AppleScriptSourceParser(normalized);
	return parser.parse();
}

async function runAppleScriptParsed<T extends AppleScriptValue>(script: string): Promise<T> {
	const rawResult = await runAppleScript(script, { humanReadableOutput: false });
	return parseAppleScriptResult(rawResult) as T;
}

type AppleScriptRecord = Record<string, unknown>;

type Note = {
	name: string;
	content: string;
	creationDate?: Date;
	modificationDate?: Date;
	folderName?: string;
	id?: string;
};

type CreateNoteResult = {
	success: boolean;
	note?: Note;
	message?: string;
	folderName?: string;
	usedDefaultFolder?: boolean;
};

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

function toDateFromUnix(value: unknown): Date | undefined {
	const unixSeconds = toNumber(value);
	if (unixSeconds === null) {
		return undefined;
	}
	return new Date(unixSeconds * 1000);
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

function mapNoteRecords(raw: unknown): Note[] {
	const records: AppleScriptRecord[] = [];

	if (Array.isArray(raw)) {
		records.push(
			...raw.filter(
				(item): item is AppleScriptRecord => item !== null && typeof item === "object",
			),
		);
	} else if (raw && typeof raw === "object") {
		records.push(raw as AppleScriptRecord);
	}

	const mapped = records.map((record) => {
		const name = coerceString(record.name, "Untitled Note");
		const content = coerceString(record.content, "");
		const folderName = coerceString(record.folderName);
		const id = coerceString(record.id);
		const creationDate = toDateFromUnix(record.creationDate);
		const modificationDate = toDateFromUnix(record.modificationDate);

		return {
			name,
			content,
			folderName: folderName || undefined,
			id: id || undefined,
			creationDate,
			modificationDate,
		};
	});

	mapped.sort((a, b) => {
		const aTime = a.modificationDate?.getTime() ?? a.creationDate?.getTime() ?? 0;
		const bTime = b.modificationDate?.getTime() ?? b.creationDate?.getTime() ?? 0;
		return bTime - aTime;
	});

	return mapped;
}

async function checkNotesAccess(): Promise<boolean> {
	try {
		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Notes"
		return name
	end tell
end timeout`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Notes app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

async function requestNotesAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		const hasAccess = await checkNotesAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Notes access is already granted.",
			};
		}

		return {
			hasAccess: false,
			message:
				"Notes access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Notes\n2. Enable access for your terminal/application\n3. If prompted, also allow Automation access for Notes\n4. Restart your terminal and try again",
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Notes access: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function ensureNotesAccess(): Promise<void> {
	const access = await requestNotesAccess();
	if (!access.hasAccess) {
		throw new Error(access.message);
	}
}

interface NoteFetchOptions {
	limit?: number;
	folderName?: string;
	searchText?: string;
}

async function fetchNotes({ limit, folderName, searchText }: NoteFetchOptions = {}): Promise<Note[]> {
	await ensureNotesAccess();

	const maxNotes = Math.max(1, Math.min(limit ?? CONFIG.MAX_NOTES, CONFIG.MAX_NOTES));
	const sanitizedFolder = folderName?.trim();
	const sanitizedSearch = searchText?.trim();

	const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Notes"
		set maxNotes to ${maxNotes}
		set noteList to {}
		set noteCount to 0

		set searchMode to ${sanitizedSearch ? "true" : "false"}
		set searchLower to ${sanitizedSearch ? JSON.stringify(sanitizedSearch.toLowerCase()) : "\"\""}

		set folderMode to ${sanitizedFolder ? "true" : "false"}
		set folderLower to ${sanitizedFolder ? JSON.stringify(sanitizedFolder.toLowerCase()) : "\"\""}

		set candidateFolders to {}
		if folderMode then
			repeat with folderItem in folders
				try
					set folderNameValue to name of folderItem as string
					set folderNameLower to do shell script "echo " & quoted form of folderNameValue & " | tr '[:upper:]' '[:lower:]'"
					if folderNameLower contains folderLower or folderLower contains folderNameLower then
						set end of candidateFolders to folderItem
					end if
				on error
					-- Ignore folder access errors
				end try
			end repeat
		else
			set candidateFolders to folders
		end if

		if folderMode and (count of candidateFolders) = 0 then
			return {status:"error", reason:"folder_not_found"}
		end if

		repeat with currentFolder in candidateFolders
			if noteCount >= maxNotes then exit repeat
			try
				set folderNameValue to name of currentFolder as string
				set folderNotes to notes of currentFolder

				repeat with noteItem in folderNotes
					if noteCount >= maxNotes then exit repeat
					try
						set includeNote to true
						if searchMode then
							set combinedText to ""
							try
								set combinedText to name of noteItem as string
							end try
							try
								set combinedText to combinedText & " " & plaintext of noteItem as string
							end try
							set combinedLower to do shell script "echo " & quoted form of combinedText & " | tr '[:upper:]' '[:lower:]'"
							if combinedLower does not contain searchLower then
								set includeNote to false
							end if
						end if

						if includeNote then
							set recordValue to my buildNoteRecord(noteItem, folderNameValue, ${CONFIG.MAX_CONTENT_PREVIEW})
							set end of noteList to recordValue
							set noteCount to noteCount + 1
						end if
					on error
						-- Skip problematic notes
					end try
				end repeat
			on error
				-- Skip folders we cannot access
			end try
		end repeat

		return noteList
	end tell
end timeout
${NOTE_SCRIPT_HELPERS}`;

	const rawResult = await runAppleScriptParsed(script);

	if (rawResult && typeof rawResult === "object" && !Array.isArray(rawResult)) {
		const status = coerceString((rawResult as AppleScriptRecord).status);
		if (status === "error") {
			const reason = coerceString((rawResult as AppleScriptRecord).reason);
			if (reason === "folder_not_found") {
				throw new Error(`Could not find a Notes folder named "${sanitizedFolder}".`);
			}
			throw new Error(reason || "Notes operation failed.");
		}
	}

	const notes = mapNoteRecords(rawResult);
	return notes.slice(0, maxNotes);
}

async function getAllNotes(): Promise<Note[]> {
	return fetchNotes();
}

async function findNote(searchText: string): Promise<Note[]> {
	const trimmed = searchText.trim();
	if (!trimmed) {
		return [];
	}

	return fetchNotes({ searchText: trimmed });
}

async function createNote(
	title: string,
	body: string,
	folderName: string = "Claude",
): Promise<CreateNoteResult> {
	try {
		await ensureNotesAccess();

		const trimmedTitle = title?.trim();
		if (!trimmedTitle) {
			return { success: false, message: "Note title cannot be empty" };
		}

		const formattedBody = body?.trim() ?? "";
		const tmpFile = join(tmpdir(), `apple-mcp-note-${randomUUID()}.txt`);

		await writeFile(tmpFile, formattedBody, "utf8");

		const sanitizedFolder = folderName?.trim();

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Notes"
		set noteTitle to ${JSON.stringify(trimmedTitle)}
		set targetFolderName to ${sanitizedFolder ? JSON.stringify(sanitizedFolder) : "\"\""}
		set folderMode to ${sanitizedFolder ? "true" : "false"}
		set targetFolderLower to ${sanitizedFolder ? JSON.stringify(sanitizedFolder.toLowerCase()) : "\"\""}

		set targetFolder to missing value
		if folderMode then
			repeat with folderItem in folders
				try
					set folderNameValue to name of folderItem as string
					set folderLower to do shell script "echo " & quoted form of folderNameValue & " | tr '[:upper:]' '[:lower:]'"
					if folderLower contains targetFolderLower or targetFolderLower contains folderLower then
						set targetFolder to folderItem
						exit repeat
					end if
				on error
					-- Ignore folder lookup errors
				end try
			end repeat
		end if

		set noteContent to read file POSIX file "${tmpFile}" as «class utf8»
		set usedDefault to false
		set folderNameValue to ""

		if targetFolder is missing value then
			set usedDefault to true
			set newNote to make new note with properties {name:noteTitle, body:noteContent}
			set folderNameValue to "Notes"
		else
			set newNote to make new note at targetFolder with properties {name:noteTitle, body:noteContent}
			set folderNameValue to name of targetFolder as string
		end if

		return {status:"success", noteRecord:my buildNoteRecord(newNote, folderNameValue, ${CONFIG.MAX_CONTENT_PREVIEW}), usedDefault:usedDefault}
	end tell
end timeout
${NOTE_SCRIPT_HELPERS}`;

		const rawResult = await runAppleScriptParsed(script);

		try {
			await unlink(tmpFile);
		} catch (cleanupError) {
			console.error("Failed to remove temporary note file:", cleanupError);
		}

		if (!rawResult || typeof rawResult !== "object" || Array.isArray(rawResult)) {
			return { success: false, message: "Failed to create note: no response from Notes." };
		}

		const status = coerceString((rawResult as AppleScriptRecord).status);
		if (status !== "success") {
			const reason = coerceString((rawResult as AppleScriptRecord).reason);
			return { success: false, message: reason || "Failed to create note." };
		}

		const noteRecord = (rawResult as AppleScriptRecord).noteRecord;
		const mappedNotes = mapNoteRecords(noteRecord);
		const createdNote = mappedNotes[0];
		const folderUsed = createdNote?.folderName || sanitizedFolder || "Notes";
		const usedDefaultFolder = coerceBoolean((rawResult as AppleScriptRecord).usedDefault);

		return {
			success: true,
			note: createdNote,
			folderName: folderUsed,
			usedDefaultFolder,
		};
	} catch (error) {
		return {
			success: false,
			message: `Failed to create note: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function getNotesFromFolder(
	folderName: string,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		const notes = await fetchNotes({ folderName });
		return { success: true, notes };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return { success: false, message };
	}
}

async function getRecentNotesFromFolder(
	folderName: string,
	limit = 5,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		const notes = await fetchNotes({ folderName, limit });
		return { success: true, notes };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return { success: false, message };
	}
}

async function getNotesByDateRange(
	folderName: string,
	fromDate?: string,
	toDate?: string,
	limit = 20,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		const notes = await fetchNotes({ folderName, limit: CONFIG.MAX_NOTES });
		const parsedFrom = parseDateInput(fromDate);
		const parsedTo = parseDateInput(toDate);
		const fromMs = parsedFrom?.getTime();
		const toMs = parsedTo?.getTime();

		const filtered = notes.filter((note) => {
			const reference = note.creationDate ?? note.modificationDate;
			if (!reference) {
				return true;
			}
			const timestamp = reference.getTime();
			if (fromMs !== undefined && timestamp < fromMs) {
				return false;
			}
			if (toMs !== undefined && timestamp > toMs) {
				return false;
			}
			return true;
		});

		return {
			success: true,
			notes: filtered.slice(0, Math.max(1, Math.min(limit, CONFIG.MAX_NOTES))),
		};
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return { success: false, message };
	}
}

export default {
	getAllNotes,
	findNote,
	createNote,
	getNotesFromFolder,
	getRecentNotesFromFolder,
	getNotesByDateRange,
	requestNotesAccess,
};
