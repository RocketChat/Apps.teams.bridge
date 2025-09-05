import { shortnameToUnicode } from "emojione";

export type Blockquote = {
    type: "BLOCKQUOTE";
    value: Paragraph[];
};

export type OrderedList = {
    type: "ORDERED_LIST";
    value: ListItem[];
};

export type UnorderedList = {
    type: "UNORDERED_LIST";
    value: ListItem[];
};

export type ListItem = {
    type: "LIST_ITEM";
    value: Inlines[];
    number?: number;
};

export type Tasks = {
    type: "TASKS";
    value: Task[];
};
export type Task = {
    type: "TASK";
    status: boolean;
    value: Inlines[];
};

export type CodeLine = {
    type: "CODE_LINE";
    value: Plain;
};

export type Color = {
    type: "COLOR";
    value: {
        r: number;
        g: number;
        b: number;
        a: number;
    };
};

export type BigEmoji = {
    type: "BIG_EMOJI";
    value: [Emoji] | [Emoji, Emoji] | [Emoji, Emoji, Emoji];
};

export type Emoji =
    | {
          type: "EMOJI";
          value: Plain;
          shortCode: string;
      }
    | {
          type: "EMOJI";
          value: undefined;
          unicode: string;
      };

export type Code = {
    type: "CODE";
    language: string | undefined;
    value: CodeLine[];
};

export type InlineCode = {
    type: "INLINE_CODE";
    value: Plain;
};

export type Heading = {
    type: "HEADING";
    level: 1 | 2 | 3 | 4;
    value: Plain[];
};

export type Quote = {
    type: "QUOTE";
    value: Paragraph[];
};

export type Markup = Italic | Strike | Bold | Plain | ChannelMention;
export type MarkupExcluding<T extends Markup> = Exclude<Markup, T>;

export type Bold = {
    type: "BOLD";
    value: Array<
        | MarkupExcluding<Bold>
        | Link
        | Emoji
        | UserMention
        | ChannelMention
        | InlineCode
    >;
};

export type Italic = {
    type: "ITALIC";
    value: Array<
        | MarkupExcluding<Italic>
        | Link
        | Emoji
        | UserMention
        | ChannelMention
        | InlineCode
    >;
};

export type Strike = {
    type: "STRIKE";
    value: Array<
        | MarkupExcluding<Strike>
        | Link
        | Emoji
        | UserMention
        | ChannelMention
        | InlineCode
        | Italic
        | Timestamp
    >;
};

export type Plain = {
    type: "PLAIN_TEXT";
    value: string;
};

export type LineBreak = {
    type: "LINE_BREAK";
    value: undefined;
};

export type KaTeX = {
    type: "KATEX";
    value: string;
};

export type InlineKaTeX = {
    type: "INLINE_KATEX";
    value: string;
};

export type Paragraph = {
    type: "PARAGRAPH";
    value: Array<Exclude<Inlines, Paragraph>>;
};

export type Image = {
    type: "IMAGE";
    value: {
        src: Plain;
        label: Markup;
    };
};

export type Link = {
    type: "LINK";
    value: {
        src: Plain;
        label: Markup | Markup[];
    };
};

export type UserMention = {
    type: "MENTION_USER";
    value: Plain;
};

export type ChannelMention = {
    type: "MENTION_CHANNEL";
    value: Plain;
};

export type Timestamp = {
    type: "TIMESTAMP";
    value: {
        timestamp: string;
        format: "t" | "T" | "d" | "D" | "f" | "F" | "R";
    };
    fallback?: Plain;
};

export type Types = {
    BOLD: Bold;
    PARAGRAPH: Paragraph;
    PLAIN_TEXT: Plain;
    ITALIC: Italic;
    STRIKE: Strike;
    CODE: Code;
    CODE_LINE: CodeLine;
    INLINE_CODE: InlineCode;
    HEADING: Heading;
    QUOTE: Quote;
    LINK: Link;
    MENTION_USER: UserMention;
    MENTION_CHANNEL: ChannelMention;
    EMOJI: Emoji;
    BIG_EMOJI: BigEmoji;
    COLOR: Color;
    TASKS: Tasks;
    TASK: Task;
    UNORDERED_LIST: UnorderedList;
    ORDERED_LIST: OrderedList;
    LIST_ITEM: ListItem;
    IMAGE: Image;
    LINE_BREAK: LineBreak;
};

export type ASTNode =
    | BigEmoji
    | Bold
    | Paragraph
    | Plain
    | Italic
    | Strike
    | Code
    | CodeLine
    | InlineCode
    | Heading
    | Quote
    | Link
    | UserMention
    | ChannelMention
    | Emoji
    | Color
    | Tasks;

export type TypesKeys = keyof Types;

export type Inlines =
    | Markup
    | Timestamp
    | InlineCode
    | Image
    | Link
    | UserMention
    | Emoji
    | Color
    | InlineKaTeX;

export type Blocks =
    | Code
    | Heading
    | Quote
    | ListItem
    | Tasks
    | OrderedList
    | UnorderedList
    | LineBreak
    | KaTeX;

type MessageBlock =
    | Emoji
    | ChannelMention
    | UserMention
    | Link
    | MarkupExcluding<Bold>
    | InlineCode;

type RootNode = Paragraph | Blocks | BigEmoji

export type Root = RootNode[];
type AnyNode = Root[number] | Inlines | { type: undefined; fallback: Plain };

const esc = (s: string) =>
    s.replace(
        /[&<>"']/g,
        (c) =>
            ({
                "&": "&amp;",
                "<": "&lt;",
                ">": "&gt;",
                '"': "&quot;",
                "'": "&#39;",
            }[c]!)
    );

const stripHtml = (s: string) => s.replace(/<[^>]*>/g, "");

export const createTeamsHTMLMessage = (root: Root, siteUrl: string) => {
    const renderChildren = (arr: AnyNode[] | undefined) =>
        (arr ?? []).map(render).join("");
    const renderInlineArray = (
        arr: (Inlines | { type: undefined; fallback: Plain })[]
    ) => arr.map(render).join("");

    const render = (node: AnyNode): string => {
        const n = node;
        if (!n || !n.type) {
            if (n?.fallback) return render(n.fallback);
            return "";
        }

        switch (n.type) {
            // --- blocks & groups ---
            case "PARAGRAPH": {
                return `<p>${n.value.map(render).join("")}</p>`;
            }
            case "HEADING": {
                const tag = `h${n.level}`;
                return `<${tag}>${n.value.map(render).join("")}</${tag}>`;
            }
            case "UNORDERED_LIST": {
                return `<ul>${n.value
                    .map((li) => `<li>${li.value.map(render).join("")}</li>`)
                    .join("")}</ul>`;
            }
            case "ORDERED_LIST": {
                return `<ol>${n.value
                    .map((li) => {
                        const liNode = li;
                        const numberAttr =
                            liNode.number != null
                                ? ` value="${liNode.number}"`
                                : "";
                        return `<li${numberAttr}>${liNode.value
                            .map(render)
                            .join("")}</li>`;
                    })
                    .join("")}</ol>`;
            }
            case "TASKS": {
                // Teams-friendly: represent as checklist-like list
                return `<ul>${n.value
                    .map((t) => {
                        const task = t;
                        const checkbox = task.status ? "☑" : "☐";
                        return `<li>${checkbox} ${task.value
                            .map(render)
                            .join("")}</li>`;
                    })
                    .join("")}</ul>`;
            }
            case "QUOTE": {
                return `<blockquote>${n.value.map(render).join("")}</blockquote>`;
            }
            case "CODE": {
                const code = n.value
                    .map((cl) => cl.value.value)
                    .join("\n");
                const lang = n.language
                    ? ` data-lang="${esc(n.language)}" class="language-${esc(
                        n.language
                    )}"`
                    : "";
                return `<pre><code${lang}>${esc(code)}</code></pre>`;
            }
            case "KATEX": {
                // Defer actual KaTeX rendering; Teams HTML usually needs pre-rendered math.
                return `<span>${esc(n.value)}</span>`;
            }
            case "LINE_BREAK":
                return "<br />";
            case "BIG_EMOJI": {
                return n.value.map(render).join("");
            }

            // --- inlines ---
            case "PLAIN_TEXT": {
                return esc(n.value);
            }
            case "BOLD": {
                return `<strong>${renderInlineArray(n.value as any)}</strong>`;
            }
            case "ITALIC": {
                return `<em>${renderInlineArray(n.value as any)}</em>`;
            }
            case "STRIKE": {
                return `<s>${renderInlineArray(n.value as any)}</s>`;
            }
            case "INLINE_CODE": {
                return `<code>${esc(n.value.value)}</code>`;
            }
            case "LINK": {
                const href = esc(n.value.src.value);
                const label = Array.isArray(n.value.label)
                    ? n.value.label.map(render).join("")
                    : render(n.value.label as any);
                return `<a href=\"${href}\" title=\"${label}\" target=\"_blank\" rel=\"noreferrer noopener\">${label}</a>`;
            }
            case "IMAGE": {
                const src = esc(n.value.src.value);
                const alt = esc(
                    stripHtml(
                        Array.isArray(n.value.label)
                            ? n.value.label.map(render).join("")
                            : render(n.value.label as any)
                    )
                );
                return `<img src="${src}" alt="${alt}" />`;
            }
            case "MENTION_USER": {
                return `@${esc(n.value.value)}`;
            }
            case "MENTION_CHANNEL": {
                return `#${esc(n.value.value)}`;
            }
            case "EMOJI": {
                const v = 'shortCode' in n ? shortnameToUnicode(`:${n.shortCode}:`) : n.unicode;
                return esc(v);
            }
            case "COLOR": {
                const { r, g, b, a } = n.value; // a is 0..255 per your type
                const alpha = (a ?? 255) / 255; // CSS expects 0..1
                const swatch = `<span style="background-color: rgba(${r}, ${g}, ${b}, ${alpha}); display:inline-block; width:1em; height:1em; vertical-align:middle; margin-inline-end:.5em;"></span>`;
                return `${swatch}rgba(${r}, ${g}, ${b}, ${alpha})`;
            }
            case "INLINE_KATEX": {
                return `<span>${esc(n.value)}</span>`;
            }
            case "TIMESTAMP": {
                // Teams doesn't render Discord-like formats; print fallback if present
                if (n.fallback) return esc(n.fallback.value);
                return `<time data-timestamp="${esc(
                    n.value.timestamp
                )}" data-format="${n.value.format}">${esc(
                    n.value.timestamp
                )}</time>`;
            }

            default:
                return "";
        }
    };
    return root.map(render).join("");
};


export const attachAttachments = (html: string, attachmentIds: string[]) => {
    return (html + `<p>${attachmentIds.map(id => `<attachment id="${id}"></attachment>`).join("")}</p>`);
}
