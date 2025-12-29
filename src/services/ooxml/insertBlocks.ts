/**
 * insertBlocks - PORTED FROM LEGACY
 * Handles INSERT operations with proper styling and track changes
 * Source: taskpane.legacy.tsx / vibe-legal-beta-0.2.yaml
 */

import { log, logWarn } from '../../utils/logger';

// Declare global types
declare var DOMParser: any;
declare var XMLSerializer: any;
declare var Word: any;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface StyleToken {
    token: string;
    styleId: string;
    styleName: string;
}

interface TextSegment {
    text: string;
    bold: boolean;
    italic: boolean;
}

/**
 * Inserts blocks of content as new paragraphs, applying styles and track changes.
 * 
 * listInfo can be:
 *   - { numId, ilvl } - full list info (when known)
 *   - { level } - simple level, will extract numId from reference paragraph
 */
export async function insertBlocks(
    context: Word.RequestContext,
    blocks: any[],
    styleCache: Map<string, StyleToken>,
    insertLocation: Word.Paragraph,
    analysis?: any,
    listInfo?: { numId?: string; ilvl?: string; level?: number } | null,
    author: string = "Vibe AI"
): Promise<number> {
    log("[insertBlocks] Starting with " + blocks.length + " blocks, listInfo:", listInfo);

    let lastInsert = insertLocation;
    let insertedCount = 0;

    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    context.trackedObjects.add(lastInsert);

    const insertLocationOxmlResult = insertLocation.getOoxml();
    await context.sync();

    let docRFonts: Node | null = null;
    let docSz: Node | null = null;
    let docSzCs: Node | null = null;

    // Variables for list info extraction
    let effectiveNumId: string | null = null;
    let effectiveIlvl: string | null = null;

    try {
        const locationXml = parser.parseFromString(insertLocationOxmlResult.value, "text/xml");
        const locationRuns = locationXml.getElementsByTagNameNS(WORD_NS, "r");
        if (locationRuns.length > 0) {
            const firstRun = locationRuns[0] as Element;
            const rPr = firstRun.getElementsByTagNameNS(WORD_NS, "rPr")[0];
            if (rPr) {
                docRFonts = rPr.getElementsByTagNameNS(WORD_NS, "rFonts")[0] || null;
                docSz = rPr.getElementsByTagNameNS(WORD_NS, "sz")[0] || null;
                docSzCs = rPr.getElementsByTagNameNS(WORD_NS, "szCs")[0] || null;
            }
        }

        // Extract numId from reference paragraph if we have listInfo with level but no numId
        if (listInfo) {
            if (listInfo.numId && listInfo.ilvl) {
                // Full list info provided
                effectiveNumId = listInfo.numId;
                effectiveIlvl = listInfo.ilvl;
            } else if (listInfo.level !== undefined) {
                // Only level provided - extract numId from reference paragraph
                effectiveIlvl = String(listInfo.level);

                const pPr = locationXml.getElementsByTagNameNS(WORD_NS, "pPr")[0] ||
                    locationXml.getElementsByTagName("w:pPr")[0];
                if (pPr) {
                    const numPr = pPr.getElementsByTagNameNS(WORD_NS, "numPr")[0] ||
                        pPr.getElementsByTagName("w:numPr")[0];
                    if (numPr) {
                        const numIdEl = numPr.getElementsByTagNameNS(WORD_NS, "numId")[0] ||
                            numPr.getElementsByTagName("w:numId")[0];
                        if (numIdEl) {
                            effectiveNumId = numIdEl.getAttribute("w:val") || numIdEl.getAttribute("val") || null;
                        }
                    }
                }
                log("[insertBlocks] Extracted numId from reference:", effectiveNumId, "ilvl:", effectiveIlvl);
            }
        }
    } catch (e) {
        logWarn(" Could not extract font/list from insertion location:", e);
    }

    const insertedParagraphs: Word.Paragraph[] = [];
    const blockMetadata: { styleToken: string, content: string, originalContent: string }[] = [];

    for (const block of blocks) {
        const styleToken = block.styleId;
        const content = block.content;

        if (!content) continue;

        const tokenData = styleCache.get(styleToken);
        const targetStyleName = tokenData ? tokenData.styleName : "Normal";
        const cleanContent = stripMarkdownFormatting(content);

        const newP = lastInsert.insertParagraph(cleanContent, "After");
        newP.style = targetStyleName;

        context.trackedObjects.add(newP);
        insertedParagraphs.push(newP);
        blockMetadata.push({ styleToken, content: cleanContent, originalContent: content });

        lastInsert = newP;
        insertedCount++;
    }

    await context.sync();

    const oxmlResults: OfficeExtension.ClientResult<string>[] = [];
    for (const p of insertedParagraphs) {
        oxmlResults.push(p.getOoxml());
    }

    await context.sync();

    for (let i = 0; i < insertedParagraphs.length; i++) {
        const p = insertedParagraphs[i];
        const originalOxml = oxmlResults[i].value;
        const meta = blockMetadata[i];

        try {
            const insId = Math.floor(Math.random() * 10000000).toString();
            const date = new Date().toISOString();

            const xmlDoc = parser.parseFromString(originalOxml, "text/xml");
            const pNode = xmlDoc.getElementsByTagNameNS(WORD_NS, "p")[0] || xmlDoc.getElementsByTagName("w:p")[0];

            if (pNode) {
                let pPr = pNode.getElementsByTagNameNS(WORD_NS, "pPr")[0] || pNode.getElementsByTagName("w:pPr")[0];

                if (!pPr && meta.styleToken) {
                    const tokenData = styleCache.get(meta.styleToken);
                    if (tokenData && tokenData.styleId !== "Normal") {
                        pPr = xmlDoc.createElementNS(WORD_NS, "w:pPr");
                        const pStyle = xmlDoc.createElementNS(WORD_NS, "w:pStyle");
                        pStyle.setAttribute("w:val", tokenData.styleId);
                        pPr.appendChild(pStyle);

                        if (pNode.firstChild) {
                            pNode.insertBefore(pPr, pNode.firstChild);
                        } else {
                            pNode.appendChild(pPr);
                        }
                    }
                }

                if (pPr) {
                    const pPr_rPr = pPr.getElementsByTagNameNS(WORD_NS, "rPr")[0] || pPr.getElementsByTagName("w:rPr")[0];
                    if (pPr_rPr) {
                        const rFonts = pPr_rPr.getElementsByTagNameNS(WORD_NS, "rFonts")[0] || pPr_rPr.getElementsByTagName("w:rFonts")[0];
                        if (rFonts) {
                            pPr_rPr.removeChild(rFonts);
                        }
                    }
                }


                const existingRuns = Array.from(pNode.childNodes).filter(n =>
                    (n as Element).localName === "r" || (n as Node).nodeName.endsWith(":r")
                );
                existingRuns.forEach(run => pNode.removeChild(run));

                const segments = parseMarkdownToSegments(meta.originalContent);
                const formattedRuns = createFormattedRuns(xmlDoc, segments, WORD_NS);

                if (docRFonts || docSz) {
                    formattedRuns.forEach(run => {
                        let rPr = run.getElementsByTagNameNS(WORD_NS, "rPr")[0] as Element;
                        if (!rPr) {
                            rPr = xmlDoc.createElementNS(WORD_NS, "w:rPr");
                            run.insertBefore(rPr, run.firstChild);
                        }
                        if (docRFonts && !rPr.getElementsByTagNameNS(WORD_NS, "rFonts")[0]) {
                            rPr.insertBefore(docRFonts.cloneNode(true), rPr.firstChild);
                        }
                        if (docSz && !rPr.getElementsByTagNameNS(WORD_NS, "sz")[0]) {
                            rPr.appendChild(docSz.cloneNode(true));
                        }
                        if (docSzCs && !rPr.getElementsByTagNameNS(WORD_NS, "szCs")[0]) {
                            rPr.appendChild(docSzCs.cloneNode(true));
                        }
                    });
                }

                const isFirstBlock = (i === 0);
                const shouldForceBold = isFirstBlock && analysis?.clauseStyle?.isBold && !meta.originalContent.includes('**');

                if (shouldForceBold) {
                    formattedRuns.forEach(run => {
                        let rPr = run.getElementsByTagNameNS(WORD_NS, "rPr")[0] as Element;
                        if (!rPr) {
                            rPr = xmlDoc.createElementNS(WORD_NS, "w:rPr");
                            run.insertBefore(rPr, run.firstChild);
                        }
                        if (!rPr.getElementsByTagNameNS(WORD_NS, "b")[0]) {
                            const b = xmlDoc.createElementNS(WORD_NS, "w:b");
                            rPr.appendChild(b);
                        }
                    });
                }

                const insNode = xmlDoc.createElementNS(WORD_NS, "w:ins");
                insNode.setAttribute("w:id", insId);
                insNode.setAttribute("w:author", author);
                insNode.setAttribute("w:date", date);

                formattedRuns.forEach(run => insNode.appendChild(run));
                pNode.appendChild(insNode);

            }

            const body = xmlDoc.getElementsByTagNameNS(WORD_NS, "body")[0] || xmlDoc.getElementsByTagName("w:body")[0];
            if (body) {
                const sectPr = body.getElementsByTagNameNS(WORD_NS, "sectPr")[0] || body.getElementsByTagName("w:sectPr")[0];
                if (sectPr) body.removeChild(sectPr);
            }

            let finalOxml = serializer.serializeToString(xmlDoc);

            // NEW: Inject numPr if inserting into a numbered list
            if (effectiveNumId && effectiveIlvl) {
                finalOxml = injectNumPrIntoOxml(finalOxml, effectiveNumId, effectiveIlvl);
                log("[insertBlocks] Injected numPr - numId:", effectiveNumId, "ilvl:", effectiveIlvl);
            }

            p.insertOoxml(finalOxml, "Replace");

        } catch (e) {
            console.error(`[Vibe] Redline Error for ${meta.styleToken}:`, e);
        }
    }

    await context.sync();

    for (const p of insertedParagraphs) {
        context.trackedObjects.remove(p);
    }
    context.trackedObjects.remove(insertLocation);

    await context.sync();

    log("[insertBlocks] Completed: " + insertedCount + " paragraphs inserted");
    return insertedCount;
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function stripMarkdownFormatting(text: string): string {
    if (!text) return "";
    return text
        .replace(/\*\*(.+?)\*\*/g, "$1")
        .replace(/\*(.+?)\*/g, "$1")
        .replace(/__(.+?)__/g, "$1")
        .replace(/`(.+?)`/g, "$1")
        .replace(/^#+\s*/gm, "")
        .replace(/^[-*]\s+/gm, "")
        .replace(/\n+/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function parseMarkdownToSegments(content: string): TextSegment[] {
    const segments: TextSegment[] = [];
    // Full regex from legacy: handles **bold**, *italic*, and plain text
    const regex = /(\*\*(.+?)\*\*)|(\*([^*]+?)\*)|([^*]+)/g;
    let match;

    while ((match = regex.exec(content)) !== null) {
        if (match[2]) {
            // Bold: **text**
            segments.push({ text: match[2], bold: true, italic: false });
        } else if (match[4]) {
            // Italic: *text*
            segments.push({ text: match[4], bold: false, italic: true });
        } else if (match[5]) {
            // Plain text
            segments.push({ text: match[5], bold: false, italic: false });
        }
    }

    return segments.length > 0 ? segments : [{ text: content, bold: false, italic: false }];
}

function createFormattedRuns(xmlDoc: Document, segments: TextSegment[], WORD_NS: string): Element[] {
    const runs: Element[] = [];

    for (const seg of segments) {
        if (!seg.text || seg.text.length === 0) continue;

        const run = xmlDoc.createElementNS(WORD_NS, "w:r");

        if (seg.bold || seg.italic) {
            const rPr = xmlDoc.createElementNS(WORD_NS, "w:rPr");
            if (seg.bold) {
                const b = xmlDoc.createElementNS(WORD_NS, "w:b");
                rPr.appendChild(b);
            }
            if (seg.italic) {
                const i = xmlDoc.createElementNS(WORD_NS, "w:i");
                rPr.appendChild(i);
            }
            run.appendChild(rPr);
        }

        const t = xmlDoc.createElementNS(WORD_NS, "w:t");
        t.setAttribute("xml:space", "preserve");
        t.textContent = seg.text;
        run.appendChild(t);

        runs.push(run);
    }

    return runs;
}

function injectNumPrIntoOxml(oxml: string, numId: string, ilvl: string): string {
    const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(oxml, "text/xml");

    let pNode = xmlDoc.getElementsByTagName("w:p")[0] as Element;
    if (!pNode) return oxml;

    let pPr = pNode.getElementsByTagName("w:pPr")[0] as Element;

    if (!pPr) {
        pPr = xmlDoc.createElementNS(NS_W, "w:pPr");
        pNode.insertBefore(pPr, pNode.firstChild);
    }

    const existingNumPr = pPr.getElementsByTagName("w:numPr")[0];
    if (existingNumPr) {
        pPr.removeChild(existingNumPr);
    }

    const numPrEl = xmlDoc.createElementNS(NS_W, "w:numPr");

    const ilvlEl = xmlDoc.createElementNS(NS_W, "w:ilvl");
    ilvlEl.setAttributeNS(NS_W, "w:val", ilvl);
    numPrEl.appendChild(ilvlEl);

    const numIdEl = xmlDoc.createElementNS(NS_W, "w:numId");
    numIdEl.setAttributeNS(NS_W, "w:val", numId);
    numPrEl.appendChild(numIdEl);

    const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
    if (pStyle && pStyle.nextSibling) {
        pPr.insertBefore(numPrEl, pStyle.nextSibling);
    } else if (pStyle) {
        pPr.appendChild(numPrEl);
    } else {
        pPr.insertBefore(numPrEl, pPr.firstChild);
    }

    return serializer.serializeToString(xmlDoc);
}
