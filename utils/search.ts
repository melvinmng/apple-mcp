import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

export type SearchResult = {
	id: string;
	title: string;
	url: string;
};

export type FetchResult = {
	id: string;
	title: string;
	text: string;
	url: string;
	metadata?: Record<string, string>;
};

type SearchDocument = {
	id: string;
	title: string;
	url: string;
	content: string;
	contentLower: string;
	titleLower: string;
};

const DEFAULT_METADATA = Object.freeze({ source: "readme" });

const MAX_RESULTS = 5;
const README_FILENAME = "README.md";
const README_REPO_URL = "https://github.com/supermemoryai/apple-mcp";

let cachedDocuments: SearchDocument[] | null = null;

function slugify(value: string): string {
	return value
		.trim()
		.toLowerCase()
		.replace(/[^a-z0-9\s-]/g, "")
		.replace(/\s+/g, "-")
		.replace(/-+/g, "-")
		.replace(/^[-]+|[-]+$/g, "");
}

async function loadReadmeSections(): Promise<SearchDocument[]> {
	try {
		const __filename = fileURLToPath(import.meta.url);
		const __dirname = path.dirname(__filename);
		const rootDir = path.resolve(__dirname, "..");
		const readmePath = path.resolve(rootDir, README_FILENAME);
		const readmeContent = await readFile(readmePath, "utf8");

		const documents: SearchDocument[] = [];
		const normalizedContent = readmeContent.replace(/```[\s\S]*?```/g, (block) =>
			block.replace(/\n/g, " "),
		);

		documents.push(
			createDocument(
				"readme-overview",
				"Project Overview",
				README_REPO_URL,
				normalizedContent,
			),
		);

		const headingRegex = /^(#{2,3})\s+(.+)$/gm;
		const matches = Array.from(readmeContent.matchAll(headingRegex));
		let sectionCounter = 0;

		for (let i = 0; i < matches.length; i++) {
			const match = matches[i];
			if (match.index === undefined) {
				continue;
			}

			const headingLevel = match[1].length;
			const title = match[2].replace(/[*`_~]/g, "").trim();
			const start = match.index + match[0].length;
			const end = i + 1 < matches.length && matches[i + 1].index !== undefined
				? matches[i + 1].index!
				: readmeContent.length;

			const sectionContent = readmeContent
				.slice(start, end)
				.replace(/#+\s+([^\n]+)/g, " $1 ")
				.replace(/[*`_~]/g, "")
				.replace(/\s+/g, " ")
				.trim();

			if (sectionContent.length === 0) {
				continue;
			}

			const slug = slugify(title);
			const url = slug
				? `${README_REPO_URL}#${slug}`
				: README_REPO_URL;

			const sectionId = headingLevel === 2
				? `readme-section-${sectionCounter}`
				: `readme-subsection-${sectionCounter}`;

			sectionCounter += 1;

			documents.push(createDocument(sectionId, title, url, sectionContent));
		}

		if (documents.length === 0) {
			return [
				createDocument(
					"readme-fallback",
					"Apple MCP Documentation",
					README_REPO_URL,
					normalizedContent,
				),
			];
		}

		return documents;
	} catch (error) {
		console.error("Failed to load README sections for search:", error);
		return [
			createDocument(
				"search-fallback",
				"Apple MCP",
				README_REPO_URL,
				"Apple MCP tools provide integrations with Contacts, Notes, Messages, Mail, Reminders, Calendar, and Maps on macOS.",
			),
		];
	}
}

function createDocument(
	id: string,
	title: string,
	url: string,
	content: string,
): SearchDocument {
	const normalizedContent = content.replace(/\s+/g, " ").trim();

	return {
		id,
		title,
		url,
		content: normalizedContent,
		contentLower: normalizedContent.toLowerCase(),
		titleLower: title.toLowerCase(),
	};
}

async function ensureDocumentsLoaded(): Promise<SearchDocument[]> {
	if (cachedDocuments) {
		return cachedDocuments;
	}

	cachedDocuments = await loadReadmeSections();
	return cachedDocuments;
}

function computeScore(doc: SearchDocument, query: string, tokens: string[]): number {
	const titleScore = doc.titleLower.includes(query) ? 6 : 0;
	const contentScore = doc.contentLower.includes(query) ? 3 : 0;

	const tokenScore = tokens.reduce((score, token) => {
		let currentScore = score;

		if (doc.titleLower.includes(token)) {
			currentScore += 3;
		}

		if (doc.contentLower.includes(token)) {
			currentScore += 1;
		}

		return currentScore;
	}, 0);

	return titleScore + contentScore + tokenScore;
}

function tokenize(query: string): string[] {
	return query
		.trim()
		.toLowerCase()
		.split(/\s+/)
		.filter((token) => token.length > 1 || /[a-z0-9]/.test(token));
}

async function search(query: string, limit: number = MAX_RESULTS): Promise<SearchResult[]> {
	const trimmedQuery = query.trim();

	if (!trimmedQuery) {
		return [];
	}

	const documents = await ensureDocumentsLoaded();
	const normalizedQuery = trimmedQuery.toLowerCase();
	const tokens = tokenize(trimmedQuery);

	const scored = documents
		.map((doc, index) => ({
			doc,
			score: computeScore(doc, normalizedQuery, tokens),
			index,
		}))
		.filter((entry) => entry.score > 0)
		.sort((a, b) => {
			if (b.score === a.score) {
				return a.index - b.index;
			}
			return b.score - a.score;
		})
		.slice(0, Math.max(1, limit));

	return scored.map(({ doc }) => ({
		id: doc.id,
		title: doc.title,
		url: doc.url,
	}));
}

async function fetchDocument(id: string): Promise<FetchResult> {
	const targetId = id.trim();

	if (!targetId) {
		throw new Error("Document id is required");
	}

	const documents = await ensureDocumentsLoaded();
	const match = documents.find((doc) => doc.id === targetId);

	if (!match) {
		throw new Error(`No document found for id '${targetId}'`);
	}

	return {
		id: match.id,
		title: match.title,
		text: match.content,
		url: match.url,
		metadata: DEFAULT_METADATA,
	};
}

export default { search, fetchDocument };
