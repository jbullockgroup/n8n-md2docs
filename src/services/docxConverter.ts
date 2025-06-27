import {
Document,
Paragraph,
TextRun,
HeadingLevel,
Packer,
AlignmentType,
Table,
TableRow,
TableCell,
BorderStyle
} from 'docx';
import { marked } from 'marked';
import { logger } from 'firebase-functions/v2';
import type { Token, Tokens } from 'marked';

const headingLevelMap = {
1: HeadingLevel.HEADING_1,
2: HeadingLevel.HEADING_2,
3: HeadingLevel.HEADING_3,
4: HeadingLevel.HEADING_4,
5: HeadingLevel.HEADING_5,
6: HeadingLevel.HEADING_6
};

// Custom paragraph break marker
const CUSTOM_PARAGRAPH_BREAK = '\\p';

function processFormattedText(text: string): TextRun[] {
const lines = text.split('\n');
const allRuns: TextRun[] = [];

for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
const line = lines[lineIndex];

const parts = line.split(/(\*\*.*?\*\*|\*.*?\*|`.*?`)/g);
const lineRuns = parts.filter(part => part !== '').map((part: string) => {
if (part.startsWith('**') && part.endsWith('**')) {
return new TextRun({ text: part.slice(2, -2), bold: true, font: 'New Times Roman', size: 24 });
} else if (part.startsWith('*') && part.endsWith('*') && !part.startsWith('**')) {
return new TextRun({ text: part.slice(1, -1), italics: true, font: 'New Times Roman', size: 24 });
} else if (part.startsWith('`') && part.endsWith('`')) {
return new TextRun({ text: part.slice(1, -1), font: 'Courier New', size: 24 });
} else {
return new TextRun({ text: part, font: 'New Times Roman', size: 24 });
}
});

allRuns.push(...lineRuns);

if (lineIndex < lines.length - 1) {
allRuns.push(new TextRun({ text: '', break: 1, font: 'New Times Roman', size: 24 }));
}
}

return allRuns;
}

export async function convertMarkdownToDocx(markdownContent: string): Promise<Buffer> {
try {
logger.info('Input markdown sample:', {
sample: markdownContent.substring(0, Math.min(200, markdownContent.length)),
length: markdownContent.length
});

// Step 1: Split content by custom paragraph breaks and process each section
const sections = markdownContent.split(CUSTOM_PARAGRAPH_BREAK);
const allChildren: Paragraph[] = [];

for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex++) {
const section = sections[sectionIndex];

// Skip empty sections
if (section.trim() === '') {
if (sectionIndex > 0) {
// Add a custom paragraph break for empty sections (but not at the beginning)
allChildren.push(new Paragraph({
children: [new TextRun({ text: '', font: 'New Times Roman', size: 24 })],
spacing: { before: 120, after: 120 }
}));
}
continue;
}

// Process this section normally with markdown
const tokens = marked.lexer(section.trim());
const children: Paragraph[] = [];
let lastTokenType: string | null = null;
let consecutiveBreaks = 0;
let listCounter = 0;

for (const token of tokens) {
if (lastTokenType && lastTokenType !== token.type) {
if (token.type === 'space') {
consecutiveBreaks++;
if (consecutiveBreaks <= 1) {
children.push(new Paragraph({ spacing: { before: 80, after: 80 } }));
}
} else {
consecutiveBreaks = 0;
}
}

switch (token.type) {
case 'heading': {
consecutiveBreaks = 0;
const headingToken = token as Tokens.Heading;
children.push(new Paragraph({
text: headingToken.text,
heading: headingLevelMap[headingToken.depth as keyof typeof headingLevelMap],
spacing: { before: 200, after: 100 }
}));
break;
}

case 'paragraph': {
consecutiveBreaks = 0;
const paragraphToken = token as Tokens.Paragraph;
const runs = processFormattedText(paragraphToken.text);
children.push(new Paragraph({
children: runs,
spacing: { before: 60, after: 60, line: 300, lineRule: 'auto' }
}));
break;
}

case 'list': {
consecutiveBreaks = 0;
const listToken = token as Tokens.List;
const numberingRef = listToken.ordered ? `list-${listCounter++}` : undefined;
let isFirstItem = true;
for (const item of listToken.items) {
const runs = processFormattedText(item.text);
children.push(new Paragraph({
children: runs,
...(listToken.ordered
? { numbering: { reference: numberingRef!, level: 0 } }
: { bullet: { level: 0 } }),
spacing: { before: isFirstItem ? 80 : 40, after: 40, line: 300, lineRule: 'auto' },
indent: { left: 720, hanging: 360 }
}));
isFirstItem = false;
}
break;
}

case 'blockquote': {
consecutiveBreaks = 0;
const blockquoteToken = token as Tokens.Blockquote;
for (const quoteToken of blockquoteToken.tokens) {
if (quoteToken.type === 'paragraph') {
const paraToken = quoteToken as Tokens.Paragraph;
const runs = processFormattedText(paraToken.text);
children.push(new Paragraph({
children: runs,
spacing: { before: 60, after: 60, line: 300, lineRule: 'auto' },
indent: { left: 720 },
border: {
left: {
color: "AAAAAA",
space: 15,
style: BorderStyle.SINGLE,
size: 15
}
}
}));
}
}
break;
}

case 'code': {
consecutiveBreaks = 0;
const codeToken = token as Tokens.Code;
const runs = processFormattedText(codeToken.text).map(run =>
new TextRun({ ...run, font: 'Courier New', size: 20 })
);
children.push(new Paragraph({
children: runs,
spacing: { before: 80, after: 80, line: 300, lineRule: 'auto' },
shading: { type: "clear", fill: "F5F5F5" }
}));
break;
}

case 'hr': {
consecutiveBreaks = 0;
children.push(new Paragraph({
border: {
bottom: {
color: "AAAAAA",
space: 1,
style: BorderStyle.SINGLE,
size: 1
}
},
spacing: { before: 120, after: 120 }
}));
break;
}

case 'table': {
consecutiveBreaks = 0;
const tableToken = token as Tokens.Table;
const rows: TableRow[] = [];

const headerCells = tableToken.header.map((cell: { text: string }) =>
new TableCell({
children: [new Paragraph({
children: processFormattedText(cell.text),
spacing: { before: 40, after: 40 }
})],
shading: { fill: "EEEEEE" }
})
);
rows.push(new TableRow({ children: headerCells }));

for (const row of tableToken.rows) {
const rowCells = row.map((cell: { text: string }) =>
new TableCell({
children: [new Paragraph({
children: processFormattedText(cell.text),
spacing: { before: 40, after: 40 }
})]
})
);
rows.push(new TableRow({ children: rowCells }));
}

const table = new Table({
rows,
width: { size: 100, type: "pct" }
});

children.push(new Paragraph({ children: [table] }));
break;
}

case 'space':
break;

default:
logger.info(`Unhandled token type: ${token.type}`);
consecutiveBreaks = 0;
break;
}

lastTokenType = token.type;
}

// Add this section's paragraphs to the overall document
allChildren.push(...children);

// Add custom paragraph break between sections (except after the last section)
if (sectionIndex < sections.length - 1) {
allChildren.push(new Paragraph({
children: [new TextRun({ text: '', font: 'New Times Roman', size: 24 })],
spacing: { before: 120, after: 120 }
}));
}
}

const numberingConfigs = [];
for (let i = 0; i < 100; i++) { // Arbitrary high number to handle all lists
numberingConfigs.push({
reference: `list-${i}`,
levels: [{
level: 0,
format: "decimal",
text: "%1.",
alignment: AlignmentType.LEFT,
style: {
paragraph: {
indent: { left: 720, hanging: 360 }
}
}
}]
});
}

const doc = new Document({
numbering: {
config: numberingConfigs
},
styles: {
default: {
document: {
run: { font: 'New Times Roman', size: 24 }
},
heading1: {
run: { size: 44, bold: true, color: "000000", font: 'New Times Roman' },
paragraph: { spacing: { before: 200, after: 100, line: 300, lineRule: 'auto' } }
},
heading2: {
run: { size: 36, bold: true, color: "000000", font: 'New Times Roman' },
paragraph: { spacing: { before: 160, after: 80, line: 300, lineRule: 'auto' } }
},
heading3: {
run: { size: 28, bold: true, color: "000000", font: 'New Times Roman' },
paragraph: { spacing: { before: 120, after: 60, line: 300, lineRule: 'auto' } }
}
},
paragraphStyles: [
{
id: "codeStyle",
name: "Code Style",
basedOn: "Normal",
run: { font: "Courier New", size: 20 },
paragraph: { spacing: { before: 80, after: 80, line: 300, lineRule: 'auto' } }
}
]
},
sections: [{
properties: {},
children: allChildren
}]
});

logger.info(`Generated DOCX with ${allChildren.length} paragraphs`);
return await Packer.toBuffer(doc);
} catch (error) {
logger.error('Error converting markdown to docx:', error);
throw error;
}
}
