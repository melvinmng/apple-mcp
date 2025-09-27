import { runAppleScript } from "run-applescript";
import { writeFile, unlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";

// Configuration
const CONFIG = {
	// Maximum emails to process (to avoid performance issues)
	MAX_EMAILS: 20,
	// Maximum content length for previews
	MAX_CONTENT_PREVIEW: 300,
	// Timeout for operations
	TIMEOUT_MS: 10000,
};

const APPLESCRIPT_TIMEOUT_SECONDS = Math.max(1, Math.ceil(CONFIG.TIMEOUT_MS / 1000));

const MESSAGE_HELPERS = `
on sortMessagesDescending(messageList)
	try
		set sortedMessages to sort messageList by date sent
		set reversedList to {}
		set totalCount to count of sortedMessages
		repeat with i from totalCount to 1 by -1
			set end of reversedList to item i of sortedMessages
		end repeat
		return reversedList
	on error
		return messageList
	end try
end sortMessagesDescending

on buildMessageRecord(currentMsg, mailboxNameValue)
	set emailSubject to ""
	try
		set emailSubject to subject of currentMsg as string
	on error
		set emailSubject to "(No subject)"
	end try

	set emailSender to ""
	try
		set emailSender to sender of currentMsg as string
	on error
		set emailSender to "Unknown sender"
	end try

	set emailDate to ""
	try
		set emailDate to (date sent of currentMsg) as string
	on error
		try
			set emailDate to (date received of currentMsg) as string
		on error
			set emailDate to ""
		end try
	end try

	set emailContent to "[Content not available]"
	try
		set fullContent to content of currentMsg as string
		if (length of fullContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
			set emailContent to (text 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of fullContent) & "..."
		else
			set emailContent to fullContent
		end if
	on error
		set emailContent to "[Content not available]"
	end try

	set isMessageRead to false
	try
		set isMessageRead to read status of currentMsg as boolean
	on error
		set isMessageRead to false
	end try

	return {subject:emailSubject, sender:emailSender, dateSent:emailDate, content:emailContent, isRead:isMessageRead, mailbox:mailboxNameValue}
end buildMessageRecord
`;

interface EmailMessage {
	subject: string;
	sender: string;
	dateSent: string;
	content: string;
	isRead: boolean;
	mailbox: string;
}

type AppleScriptRecord = Record<string, unknown>;

function toAppleScriptString(value: string): string {
	return JSON.stringify(value ?? "");
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

function coerceString(value: unknown, fallback = ""): string {
	if (typeof value === "string") {
		return value;
	}
	if (value === null || value === undefined) {
		return fallback;
	}
	return String(value);
}

function toIsoDate(value: unknown): string {
	const asString = coerceString(value);
	if (!asString) {
		return new Date().toISOString();
	}
	const parsed = Date.parse(asString);
	if (Number.isNaN(parsed)) {
		return asString;
	}
	return new Date(parsed).toISOString();
}

function mapEmailRecords(raw: unknown, fallbackMailbox?: string): EmailMessage[] {
	if (!Array.isArray(raw)) {
		return [];
	}

	return raw
		.filter((item): item is AppleScriptRecord =>
			item !== null && typeof item === "object",
		)
		.map((record) => {
			const subject = coerceString(record.subject, "No subject");
			const sender = coerceString(record.sender, "Unknown sender");
			const content = coerceString(record.content, "[Content not available]");
			const dateSent = toIsoDate(
				record.dateSent ?? record.date ?? record.sentDate ?? record.receivedDate,
			);
			const mailbox = coerceString(record.mailbox ?? fallbackMailbox ?? "Unknown mailbox");
			const isRead = coerceBoolean(record.isRead, !!record.read);

			return {
				subject,
				sender,
				dateSent,
				content,
				isRead,
				mailbox,
			};
		})
		.filter((email) => email.subject || email.sender);
}

async function ensureMailAccess(): Promise<void> {
	const accessResult = await requestMailAccess();
	if (!accessResult.hasAccess) {
		throw new Error(accessResult.message);
	}
}

/**
 * Check if Mail app is accessible
 */
async function checkMailAccess(): Promise<boolean> {
	try {
		const script = `
tell application "Mail"
    return name
end tell`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Mail app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

/**
 * Request Mail app access and provide instructions if not available
 */
async function requestMailAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		// First check if we already have access
		const hasAccess = await checkMailAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Mail access is already granted."
			};
		}

		// If no access, provide clear instructions
		return {
			hasAccess: false,
			message: "Mail access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Automation\n2. Find your terminal/app in the list and enable 'Mail'\n3. Make sure Mail app is running and configured with at least one account\n4. Restart your terminal and try again"
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Mail access: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Get unread emails from Mail app (limited for performance)
 */
async function getUnreadMails(limit = 10): Promise<EmailMessage[]> {
	try {
		await ensureMailAccess();
		const maxEmails = Math.min(limit, CONFIG.MAX_EMAILS);
		if (maxEmails <= 0) {
			return [];
		}

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Mail"
		set emailList to {}
		set emailCount to 0
		set desiredLimit to ${maxEmails}

		repeat with currentAccount in accounts
			if emailCount >= desiredLimit then exit repeat
			try
				set inboxMailbox to inbox of currentAccount
				set unreadMessages to (messages of inboxMailbox whose read status is false)
				set sortedMessages to my sortMessagesDescending(unreadMessages)
				repeat with currentMsg in sortedMessages
					if emailCount >= desiredLimit then exit repeat
					try
						set recordData to my buildMessageRecord(currentMsg, name of inboxMailbox as string)
						set emailCount to emailCount + 1
						set end of emailList to recordData
					on error
						-- Skip messages we cannot read
					end try
				end repeat
			on error
				-- Skip accounts we cannot read
			end try
		end repeat

		return emailList
	end tell
end timeout
${MESSAGE_HELPERS}`;

		const rawResult = await runAppleScript(script);
		return mapEmailRecords(rawResult);
	} catch (error) {
		console.error(
			`Error getting unread emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw error;
	}
}

/**
 * Search for emails by search term
 */
async function searchMails(
	searchTerm: string,
	limit = 10,
): Promise<EmailMessage[]> {
	try {
		await ensureMailAccess();

		if (!searchTerm || searchTerm.trim() === "") {
			return [];
		}

		const maxEmails = Math.min(limit, CONFIG.MAX_EMAILS);
		const cleanSearchTerm = searchTerm.trim();

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Mail"
		set emailList to {}
		set emailCount to 0
		set desiredLimit to ${maxEmails}
		set searchValue to ${toAppleScriptString(cleanSearchTerm)}
		set searchLower to do shell script "echo " & quoted form of searchValue & " | tr '[:upper:]' '[:lower:]'"

		repeat with currentAccount in accounts
			if emailCount >= desiredLimit then exit repeat
			try
				set inboxMailbox to inbox of currentAccount
				set allMessages to messages of inboxMailbox
				set sortedMessages to my sortMessagesDescending(allMessages)
				repeat with currentMsg in sortedMessages
					if emailCount >= desiredLimit then exit repeat
					try
						set combinedText to ""
						try
							set combinedText to subject of currentMsg as string
						on error
							set combinedText to ""
						end try
						try
							set combinedText to combinedText & " " & (content of currentMsg as string)
						on error
							-- Ignore content fetch errors for filtering
						end try

						set combinedLower to do shell script "echo " & quoted form of combinedText & " | tr '[:upper:]' '[:lower:]'"
						if combinedLower contains searchLower then
							set recordData to my buildMessageRecord(currentMsg, name of inboxMailbox as string)
							set end of emailList to recordData
							set emailCount to emailCount + 1
						end if
					on error
						-- Skip messages we cannot inspect
					end try
				end repeat
			on error
				-- Skip accounts without inbox access
			end try
		end repeat

		return emailList
	end tell
end timeout
${MESSAGE_HELPERS}`;

		const rawResult = await runAppleScript(script);
		return mapEmailRecords(rawResult);
	} catch (error) {
		console.error(
			`Error searching emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw error;
	}
}

/**
 * Send an email
 */
async function sendMail(
	to: string,
	subject: string,
	body: string,
	cc?: string,
	bcc?: string,
): Promise<string | undefined> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		// Validate inputs
		if (!to || !to.trim()) {
			throw new Error("To address is required");
		}
		if (!subject || !subject.trim()) {
			throw new Error("Subject is required");
		}
		if (!body || !body.trim()) {
			throw new Error("Email body is required");
		}

		// Use file-based approach for email body to avoid AppleScript escaping issues
		const tmpFile = join(tmpdir(), `apple-mcp-mail-${randomUUID()}.txt`);
		await writeFile(tmpFile, body.trim(), "utf8");

		const script = `
tell application "Mail"
    activate

    -- Read email body from file to preserve formatting
    set emailBody to read file POSIX file "${tmpFile}" as «class utf8»

    -- Create new message
    set newMessage to make new outgoing message with properties {subject:${toAppleScriptString(subject)}, content:emailBody, visible:true}

    tell newMessage
        make new to recipient with properties {address:${toAppleScriptString(to)}}
        ${cc ? `make new cc recipient with properties {address:${toAppleScriptString(cc)}}` : ""}
        ${bcc ? `make new bcc recipient with properties {address:${toAppleScriptString(bcc)}}` : ""}
    end tell

    send newMessage
    return "SUCCESS"
end tell`;

		const result = (await runAppleScript(script)) as string;

		// Clean up temporary file
		try {
			await unlink(tmpFile);
		} catch (e) {
			// Ignore cleanup errors
		}

		if (result === "SUCCESS") {
			return `Email sent to ${to} with subject "${subject}"`;
		} else {
			throw new Error("Failed to send email");
		}
	} catch (error) {
		console.error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw new Error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

/**
 * Get list of mailboxes (simplified for performance)
 */
async function getMailboxes(): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	    tell application "Mail"
        set mailboxNames to {}
        repeat with currentMailbox in mailboxes
            try
                set end of mailboxNames to name of currentMailbox as string
            end try
        end repeat
        return mailboxNames
    end tell
end timeout`;

		const result = await runAppleScript(script);
		if (!Array.isArray(result)) {
			return [];
		}

		return result.filter((name): name is string => typeof name === "string" && name.length > 0);
	} catch (error) {
		console.error(
			`Error getting mailboxes: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw error;
	}
}

/**
 * Get list of email accounts (simplified for performance)
 */
async function getAccounts(): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	    tell application "Mail"
        try
            set accountNames to {}
            repeat with currentAccount in accounts
                try
                    set end of accountNames to name of currentAccount as string
                end try
            end repeat
            return accountNames
        on error errMsg
            return {"Error:" & errMsg}
        end try
    end tell
end timeout`;

		const result = await runAppleScript(script);
		if (Array.isArray(result)) {
			return result
				.filter((name): name is string => typeof name === "string")
				.filter((name) => !name.startsWith("Error:"));
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting accounts: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw error;
	}
}

/**
 * Get mailboxes for a specific account
 */
async function getMailboxesForAccount(accountName: string): Promise<string[]> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		if (!accountName || !accountName.trim()) {
			return [];
		}

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Mail"
    set boxList to {}

    try
        -- Find the account
        set targetAccount to first account whose name is "${accountName.replace(/"/g, '\\"')}"
        set accountMailboxes to mailboxes of targetAccount

        repeat with i from 1 to (count of accountMailboxes)
            try
                set currentMailbox to item i of accountMailboxes
                set mailboxName to name of currentMailbox
                set boxList to boxList & {mailboxName}
            on error
                -- Skip problematic mailboxes
            end try
        end repeat
    on error
        -- Account not found or other error
        return {}
    end try

    return boxList
end tell
end timeout`;

		const result = await runAppleScript(script);
		if (!Array.isArray(result)) {
			return [];
		}

		return result.filter((name): name is string => typeof name === "string" && name.length > 0);
	} catch (error) {
		console.error(
			`Error getting mailboxes for account: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw error;
	}
}

/**
 * Get latest emails from a specific account
 */
async function getLatestMails(
	account: string,
	limit = 5,
): Promise<EmailMessage[]> {
	try {
		await ensureMailAccess();
		if (!account || !account.trim()) {
			return [];
		}

		const sanitizedAccount = account.replace(/"/g, '\\"');
		const desiredLimit = Math.min(Math.max(limit, 1), CONFIG.MAX_EMAILS);

		const script = `
with timeout of ${APPLESCRIPT_TIMEOUT_SECONDS} seconds
	tell application "Mail"
		set resultList to {}
		try
			set targetAccount to first account whose name is "${sanitizedAccount}"
		on error errMsg
			return {status:"error", reason:errMsg}
		end try

		try
			set inboxMailbox to inbox of targetAccount
		on error errMsg
			return {status:"error", reason:errMsg}
		end try

		set sortedMessages to my sortMessagesDescending(messages of inboxMailbox)
		set msgLimit to ${desiredLimit}
		if (count of sortedMessages) < msgLimit then
			set msgLimit to count of sortedMessages
		end if

		repeat with i from 1 to msgLimit
			try
				set currentMsg to item i of sortedMessages
				set recordData to my buildMessageRecord(currentMsg, name of inboxMailbox as string)
				set end of resultList to recordData
			on error
				-- Skip messages we cannot inspect
			end try
		end repeat

		return resultList
	end tell
end timeout
${MESSAGE_HELPERS}`;

		const rawResult = await runAppleScript(script);
		if (rawResult && typeof rawResult === "object" && !Array.isArray(rawResult)) {
			const status = coerceString((rawResult as AppleScriptRecord).status);
			if (status === "error") {
				const reason = coerceString((rawResult as AppleScriptRecord).reason);
				throw new Error(reason || "Failed to read account inbox");
			}
		}

		return mapEmailRecords(rawResult, account);
	} catch (error) {
		console.error("Error getting latest emails:", error);
		throw error;
	}
}

export default {
	getUnreadMails,
	searchMails,
	sendMail,
	getMailboxes,
	getAccounts,
	getMailboxesForAccount,
	getLatestMails,
	requestMailAccess,
};
