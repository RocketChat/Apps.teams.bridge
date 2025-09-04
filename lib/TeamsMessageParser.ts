import { parse } from 'himalaya';
import { decode } from "he";
import { Attachment } from './MicrosoftGraphApi';
import { retrieveMessageIdMappingByTeamsMessageIdAsync } from './PersistHelper';
import { IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { getRocketChatMessageUrl } from './UrlHelper';
import { TeamsAttachmentType } from './Const';

export type NodeType = "element" | "comment" | "text";

/** Positions (present only when includePositions: true) */
export interface Position {
    index: number;
    line: number;
    column: number;
}

export interface NodePosition {
    start: Position;
    end: Position;
}

/** Base node */
export interface BaseNode {
    type: NodeType;
    /** Included only when parser is configured with includePositions: true */
    position?: NodePosition;
}

/** Attribute on an element */
export interface Attribute {
    key: string;
    /**
     * For standalone attributes like `disabled`, the value is `null`.
     * Some emitters may omit the property entirely, so it's optional.
     */
    value?: string | null;
}

/** Element node: tags like <body>, <div>, <style> */
export interface ElementNode extends BaseNode {
    type: "element";
    tagName: string;
    children: Node[];
    attributes: Attribute[];
}

/** Comment node: <!-- comment --> */
export interface CommentNode extends BaseNode {
    type: "comment";
    content: string;
}

/** Text node */
export interface TextNode extends BaseNode {
    type: "text";
    content: string;
}

/** Union of all node kinds */
export type Node = ElementNode | CommentNode | TextNode;

/** Parser options (for reference) */
export interface ParseOptions {
    includePositions?: boolean;
}

/** Typical parse return: a list of top-level nodes (document or fragment) */
export type ParseResult = Node[];

export function parseHTML(input: string, options?: ParseOptions): ParseResult {
    return parse(input, options);
}

export async function buildRocketChatMessageText({
    nodes,
    attachments,
    read,
}: {
    nodes: ParseResult;
    attachments?: Attachment[];
    read: IRead;
}): Promise<string> {
    async function parseNode(node: Node, extra?: Record<string, any>): Promise<string> {
        if (node.type === "text") {
            return decode(node.content);
        } else if (node.type === "element") {
            switch (node.tagName) {
                case "a":
                    const href = node.attributes.find(
                        (attr) => attr.key === "href"
                    )?.value;
                    return `[${(await Promise.all(node.children.map((c) => parseNode(c)))).join("")}](${href})`;
                case "b":
                case "strong":
                    if (node.children.length === 0) {
                        return ``;
                    }
                    return `**${(await Promise.all(node.children.map((c) => parseNode(c)))).join("")}**`;
                case "i":
                case "em":
                    if (node.children.length === 0) {
                        return ``;
                    }
                    return `_${(await Promise.all(node.children.map((c) => parseNode(c)))).join("")}_`;
                case "br":
                    return `\n`;
                case "p":
                    return `${(await Promise.all(node.children.map((c) => parseNode(c)))).join("")}`;
                case "s":
                case "strike":
                    if (node.children.length === 0) {
                        return ``;
                    }
                    return `~~${(await Promise.all(node.children.map((c) => parseNode(c)))).join("")}~~`;
                case 'ul':
                    if (node.children.length === 0) {
                        return ``;
                    }
                    return (await Promise.all(node.children.map((c) => parseNode(c, { ul: true })))).join("");
                case "ol":
                    if (node.children.length === 0) {
                        return ``;
                    }
                    let index = 1;
                    return (
                        await Promise.all(
                            node.children.map((c) =>
                                parseNode(c, {
                                    ol: true,
                                    index:
                                        "tagName" in c && c.tagName === "li"
                                            ? index++
                                            : undefined,
                                })
                            )
                        )
                    ).join("");
                case "li":
                    if (node.children.length === 0) {
                        return ``;
                    }
                    if (extra?.ol) {
                        return `${extra?.index ? extra.index : 1}. ${(
                            await Promise.all(
                                node.children.map((c) => parseNode(c))
                            )
                        ).join("")}\n`;
                    } else {
                        return `- ${(
                            await Promise.all(
                                node.children.map((c) => parseNode(c))
                            )
                        ).join("")}\n`;
                    }

                case 'blockquote':
                    if (node.children.length === 0) {
                        return ``;
                    }
                    const text = (await Promise.all(node.children.map((c) => parseNode(c)))).join("");
                    const lines = text.split("\n").map(line => `> ${line}`);
                    if (lines.length > 0 && lines[lines.length - 1] === "> ") {
                        lines.pop();
                    }
                    return lines.join("\n");

                case 'emoji':
                    return node.attributes.find(attr => attr.key === 'alt')?.value ?? '';

                case "attachment":
                    const id = node.attributes.find(
                        (attr) => attr.key === "id"
                    )?.value;
                    if (!id) {
                        return ``;
                    }
                    const attachment = attachments?.find((att) => att.id === id);
                    switch (attachment?.contentType) {
                        case TeamsAttachmentType.MessageReference:
                            let msgUrl = "";
                            const mapping =
                                await retrieveMessageIdMappingByTeamsMessageIdAsync(
                                    read,
                                    attachment.id
                                );
                            if (mapping?.rocketChatMessageId) {
                                const msg = await read.getMessageReader().getById(mapping.rocketChatMessageId);
                                if (msg?.id) {
                                    msgUrl = await getRocketChatMessageUrl(read, msg.id, msg.room);
                                }
                            }
                            return msgUrl;
                        default:
                            return ``;
                    }
                default:
                    return (await Promise.all(node.children.map((c) => parseNode(c)))).join("");
            }
        } else {
            return "";
        }
    }
    const results = await Promise.all(
        nodes.map((node) => parseNode(node))
    );
    return results.join("");
}
