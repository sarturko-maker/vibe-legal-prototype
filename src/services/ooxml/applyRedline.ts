/**
 * applyRedlineToOxml - PORTED FROM LEGACY
 * DO NOT MODIFY - This is battle-tested code
 * Source: taskpane.legacy.tsx / vibe-legal-beta-0.2.yaml
 */

import { log, logWarn } from '../../utils/logger';
import { parseFormattingMarkers, stripFormattingMarkers } from '../formatting/markdownParser';

// Declare global types
declare var DOMParser: any;
declare var XMLSerializer: any;
declare var diff_match_patch: any;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

interface TextSpan {
    charStart: number;
    charEnd: number;
    textElement: Element;
    runElement: Element;
    paragraph: Element;
    container: Element;
    rPr: Element | null;
}

/**
 * Main entry point for the redlining engine.
 * Determines the best strategy (Surgical, Paragraph-Aware, or Reconstruction)
 * and applies changes to the OOXML.
 */
export function applyRedlineToOxml(
    oxml: string,
    originalText: string,
    modifiedText: string,
    author: string = "Vibe AI"
): { oxml: string; hasChanges: boolean } {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, "text/xml");
    } catch (e) {
        console.error("[Engine] Failed to parse OXML:", e);
        return { oxml, hasChanges: false };
    }

    // ENHANCED: Use robust numbering detection
    const hasNumbering = hasNumberingInOxml(xmlDoc);
    console.log('[Engine] hasNumbering detected:', hasNumbering);

    // ENHANCED: Clean up AI response for numbered lists
    const cleanModifiedText = hasNumbering
        ? sanitizeForNumberedList(originalText, modifiedText, true)
        : sanitizeAiResponse(modifiedText);

    if (cleanModifiedText.trim() === originalText.trim()) {
        return { oxml, hasChanges: false };
    }

    // Count paragraphs in OXML
    const allParagraphs = xmlDoc.getElementsByTagName("w:p");
    const oxmlParaCount = allParagraphs.length;

    // Count paragraphs in modified text - split on DOUBLE newlines (proper paragraph breaks)
    // Single \n may just be line wrapping within a paragraph
    const aiHasMultipleParagraphs = /\n\s*\n/.test(cleanModifiedText);
    const modParas = cleanModifiedText.split(/\n\s*\n/).filter((p: string) => p.trim().length > 0);
    const modParaCount = modParas.length;

    console.log('[Engine] Paragraph counts - OXML:', oxmlParaCount, 'Modified:', modParaCount);

    // PARAGRAPH-AWARE MODE DETECTION - EXPANDED
    // Use paragraph-aware mode when:
    // 1. Original has 2+ paragraphs (catches heading + body) - lowered from 3
    // 2. AI returned multi-paragraph response (has \n\n)
    // 3. Structure is changing (paragraph count differs)
    const isMultiParagraph = oxmlParaCount >= 2;  // Changed from 3 to catch heading+body
    const structureChanging = Math.abs(oxmlParaCount - modParaCount) >= 1;

    // EXPANDED: Activate paragraph-aware for any of these conditions
    const needsParagraphAware =
        (isMultiParagraph && structureChanging) ||  // Removed hasNumbering requirement
        aiHasMultipleParagraphs ||                   // NEW: AI returned \n\n paragraph breaks
        (structureChanging && hasNumbering);         // Original: numbered list structure change

    if (needsParagraphAware) {
        log("[Engine] Using PARAGRAPH-AWARE mode: " + oxmlParaCount + " original paras, " + modParaCount + " modified paras, aiMultiPara=" + aiHasMultipleParagraphs);
        return applyParagraphAwareMode(xmlDoc, originalText, cleanModifiedText, serializer, author);
    }

    // EXPANDED DETECTION: Tables, Numbered Lists, OR Complex Structures
    const tables = xmlDoc.getElementsByTagName("w:tbl");
    const hasTables = tables.length > 0;

    // Other complex structures that need preservation
    const hasBookmarks = xmlDoc.getElementsByTagName("w:bookmarkStart").length > 0;
    const hasFields = xmlDoc.getElementsByTagName("w:fldChar").length > 0;
    const hasFootnotes = xmlDoc.getElementsByTagName("w:footnoteReference").length > 0;
    const hasEndnotes = xmlDoc.getElementsByTagName("w:endnoteReference").length > 0;
    const hasComments = xmlDoc.getElementsByTagName("w:commentRangeStart").length > 0;
    const hasHyperlinks = xmlDoc.getElementsByTagName("w:hyperlink").length > 0;  // Links
    const hasSdt = xmlDoc.getElementsByTagName("w:sdt").length > 0;                // Content controls

    const needsSurgicalMode =
        hasTables ||
        hasNumbering ||
        hasBookmarks ||
        hasFields ||
        hasFootnotes ||
        hasEndnotes ||
        hasComments ||
        hasHyperlinks ||
        hasSdt;

    console.log('[Engine] Mode decision - needsSurgicalMode:', needsSurgicalMode, '(hasNumbering:', hasNumbering, 'hasTables:', hasTables, 'hasHyperlinks:', hasHyperlinks, ')');

    // Use surgical mode for any complex structure to preserve formatting
    if (needsSurgicalMode) {
        console.log('[Engine] Using SURGICAL mode');
        return applySurgicalMode(xmlDoc, originalText, cleanModifiedText, serializer, author);
    } else {
        console.log('[Engine] Using RECONSTRUCTION mode');
        return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer, author);
    }
}


// ==========================================
// MODE 1: SURGICAL MODE (Character-level diffs)
// ==========================================

function applySurgicalMode(
    xmlDoc: Document,
    originalText: string,
    modifiedText: string,
    serializer: XMLSerializer,
    author: string
): { oxml: string; hasChanges: boolean } {
    let fullText = "";
    const textSpans: TextSpan[] = [];

    // CRITICAL FIX: Only process paragraphs in the document BODY, not in styles/numbering definitions
    // getElementsByTagName returns ALL w:p elements across entire OOXML package
    const allParagraphsRaw = Array.from(xmlDoc.getElementsByTagName("w:p"));
    const allParagraphs = allParagraphsRaw.filter((p: Element) => {
        // Check if this paragraph is inside w:body (document content)
        let parent: Node | null = p.parentNode;
        while (parent) {
            const nodeName = (parent as Element).nodeName || '';
            if (nodeName === 'w:body') return true;  // In body - process it
            if (nodeName === 'w:styles' || nodeName === 'w:numbering') return false;  // In styles/numbering - skip
            parent = parent.parentNode;
        }
        return true;  // Default to processing if no clear indicator
    });
    console.log('[Engine] Filtered paragraphs:', allParagraphs.length, 'of', allParagraphsRaw.length, 'total');

    // SAFETY: Backup pPr (especially numPr) before modification
    const pPrBackup = new Map<number, Element>();
    allParagraphs.forEach((p: Element, idx) => {
        const pPr = p.getElementsByTagName("w:pPr")[0];
        if (pPr) {
            pPrBackup.set(idx, pPr.cloneNode(true) as Element);
            const hasNumPr = pPr.getElementsByTagName("w:numPr").length > 0;
            if (hasNumPr) {
                console.log('[Engine] Backup: P' + idx + ' has numPr in pPr - backed up');
            }
        }
    });
    console.log('[Engine] pPr backup complete:', pPrBackup.size, 'paragraphs backed up');

    // DEBUG: Find where numPr actually is in the OOXML
    const allNumPr = xmlDoc.getElementsByTagName("w:numPr");
    console.log('[Engine] DEBUG: Total w:numPr elements in OOXML:', allNumPr.length);
    for (let i = 0; i < Math.min(allNumPr.length, 3); i++) {
        const numPr = allNumPr[i];
        const parentChain: string[] = [];
        let parent: Node | null = numPr.parentNode;
        while (parent && parentChain.length < 5) {
            parentChain.push((parent as Element).nodeName || 'unknown');
            parent = parent.parentNode;
        }
        console.log('[Engine] DEBUG: numPr[' + i + '] parent chain:', parentChain.join(' <- '));
    }


    allParagraphs.forEach((p: Element, pIndex) => {
        const container = p.parentNode as Element;

        // PRE-PROCESSING: Normalize runs (1 run = 1 text node)
        // This ensures that processDelete can safely remove a run without affecting other text spans
        const runs = Array.from(p.getElementsByTagName("w:r"));
        for (const r of runs) {
            const textNodes = Array.from(r.getElementsByTagName("w:t"));
            if (textNodes.length > 1) {
                const rPr = r.getElementsByTagName("w:rPr")[0];
                const parent = r.parentNode;
                if (!parent) continue;

                // Keep the first text node in the original run, move others to new runs
                // Insert new runs AFTER the current run in reverse order (to keep order correct)
                for (let i = textNodes.length - 1; i > 0; i--) {
                    const t = textNodes[i];
                    const newRun = xmlDoc.createElementNS(WORD_NS, "w:r");
                    if (rPr) newRun.appendChild(rPr.cloneNode(true));
                    newRun.appendChild(t); // Moves t from r to newRun

                    if (r.nextSibling) {
                        parent.insertBefore(newRun, r.nextSibling);
                    } else {
                        parent.appendChild(newRun);
                    }
                }
            }
        }

        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === "w:r") {
                const r = child as Element;
                const rPr = r.getElementsByTagName("w:rPr")[0] || null;
                // Now guarantees at most 1 w:t per w:r
                const textNodes = r.getElementsByTagName("w:t");
                if (textNodes.length > 0) {
                    const t = textNodes[0];
                    const text = t.textContent || "";
                    if (text.length > 0) {
                        textSpans.push({ charStart: fullText.length, charEnd: fullText.length + text.length, textElement: t, runElement: r, paragraph: p, container, rPr });
                        fullText += text;
                    }
                }
            } else if (child.nodeName === "w:hyperlink") {
                Array.from(child.childNodes).forEach(hc => {
                    if (hc.nodeName === "w:r") {
                        const r = hc as Element;
                        const rPr = r.getElementsByTagName("w:rPr")[0] || null;
                        Array.from(r.childNodes).forEach(rc => {
                            if (rc.nodeName === "w:t") {
                                const t = rc as Element;
                                const text = t.textContent || "";
                                if (text.length > 0) {
                                    textSpans.push({ charStart: fullText.length, charEnd: fullText.length + text.length, textElement: t, runElement: r, paragraph: p, container, rPr });
                                    fullText += text;
                                }
                            }
                        });
                    }
                });
            }
        });
        if (pIndex < allParagraphs.length - 1) fullText += "\n";
    });

    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(fullText, modifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // Post-process to ensure whole-word replacements
    const wordCleanedDiffs = cleanupDiffsToWords(diffs);

    let currentPos = 0;
    const processedSpans = new Set();

    for (const diff of wordCleanedDiffs) {
        const [op, text] = diff;
        if (op === 0) {
            currentPos += text.length;
        } else if (op === -1) {
            processDelete(xmlDoc, textSpans, currentPos, currentPos + text.length, processedSpans, author);
            currentPos += text.length;
        } else if (op === 1) {
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            // Allow space-only insertions - only skip if completely empty
            if (textWithoutNewlines.length > 0) {
                processInsert(xmlDoc, textSpans, currentPos, textWithoutNewlines, processedSpans, author);
            }
        }
    }

    // SAFETY: Restore pPr/numPr if accidentally removed during processing
    const updatedParagraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
    updatedParagraphs.forEach((p: Element, idx) => {
        const currentPPr = p.getElementsByTagName("w:pPr")[0];
        const backupPPr = pPrBackup.get(idx);

        if (!currentPPr && backupPPr) {
            // pPr was removed - restore it
            p.insertBefore(backupPPr, p.firstChild);
        } else if (currentPPr && backupPPr) {
            // Ensure numPr wasn't stripped
            const currentNumPr = currentPPr.getElementsByTagName("w:numPr")[0];
            const backupNumPr = backupPPr.getElementsByTagName("w:numPr")[0];

            if (!currentNumPr && backupNumPr) {
                currentPPr.insertBefore(backupNumPr.cloneNode(true), currentPPr.firstChild);
            }
        }
    });

    // DEBUG: Check if track change elements were created
    const delElements = xmlDoc.getElementsByTagName("w:del");
    const insElements = xmlDoc.getElementsByTagName("w:ins");

    const result = serializer.serializeToString(xmlDoc);

    return { oxml: result, hasChanges: true };
}

function processDelete(xmlDoc: Document, textSpans: TextSpan[], startPos: number, endPos: number, processedSpans: Set<any>, author: string) {
    const affectedSpans = textSpans.filter(s => s.charEnd > startPos && s.charStart < endPos);

    for (const span of affectedSpans) {
        if (processedSpans.has(span.textElement)) continue;

        const startInSpan = Math.max(0, startPos - span.charStart);
        const endInSpan = Math.min(span.textElement.textContent?.length || 0, endPos - span.charStart);

        const deleteLen = endInSpan - startInSpan;
        if (deleteLen <= 0) continue;

        const parent = span.runElement.parentNode;
        if (!parent) continue;

        const fullText = span.textElement.textContent || "";
        const beforeText = fullText.substring(0, startInSpan);
        const deletedText = fullText.substring(startInSpan, endInSpan);
        const afterText = fullText.substring(endInSpan);

        // Entire span is deleted?
        const isWholeSpan = startPos <= span.charStart && endPos >= span.charEnd;

        if (isWholeSpan) {
            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
            parent.insertBefore(delWrapper, span.runElement);
            parent.removeChild(span.runElement);
        } else {
            if (beforeText.length > 0) {
                const beforeRun = createTextRun(xmlDoc, beforeText, span.rPr, false);
                parent.insertBefore(beforeRun, span.runElement);
            }
            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun, author);
            parent.insertBefore(delWrapper, span.runElement);
            if (afterText.length > 0) {
                const afterRun = createTextRun(xmlDoc, afterText, span.rPr, false);
                parent.insertBefore(afterRun, span.runElement);
            }
            parent.removeChild(span.runElement);
        }
        processedSpans.add(span.textElement);
    }
}

function processInsert(xmlDoc: Document, textSpans: TextSpan[], pos: number, text: string, processedSpans: Set<any>, author: string) {
    try {
        let targetSpan = textSpans.find(s => pos >= s.charStart && pos < s.charEnd);

        if (!targetSpan && pos > 0) {
            targetSpan = textSpans.find(s => pos === s.charEnd);
        }

        if (!targetSpan && pos > 0) {
            targetSpan = textSpans.find(s => s.charEnd <= pos);
            if (!targetSpan && textSpans.length > 0) {
                const before = textSpans.filter(s => s.charEnd <= pos);
                if (before.length > 0) {
                    targetSpan = before[before.length - 1];
                }
            }
        }

        if (!targetSpan && textSpans.length > 0) {
            targetSpan = textSpans[textSpans.length - 1];
        }

        if (!targetSpan) return;

        const rPr = targetSpan.rPr;
        const parent = targetSpan.runElement.parentNode;

        if (!parent) return;

        // MARKDOWN PARSING: Parse **bold**, *italic*, __underline__
        const segments = parseFormattingMarkers(text);

        const insWrapper = createTrackChange(xmlDoc, 'ins', null as any, author);

        // Remove the null child that createTrackChange added
        while (insWrapper.firstChild) insWrapper.removeChild(insWrapper.firstChild);

        for (const seg of segments) {
            const run = xmlDoc.createElementNS(WORD_NS, "w:r");

            // Clone rPr and add formatting
            const runRPr = rPr ? rPr.cloneNode(true) as Element : xmlDoc.createElementNS(WORD_NS, "w:rPr");
            if (seg.formatting.bold) {
                runRPr.appendChild(xmlDoc.createElementNS(WORD_NS, "w:b"));
            }
            if (seg.formatting.italic) {
                runRPr.appendChild(xmlDoc.createElementNS(WORD_NS, "w:i"));
            }
            if (seg.formatting.underline) {
                const u = xmlDoc.createElementNS(WORD_NS, "w:u");
                u.setAttribute("w:val", "single");
                runRPr.appendChild(u);
            }
            if (runRPr.childNodes.length > 0 || rPr) {
                run.appendChild(runRPr);
            }

            const t = xmlDoc.createElementNS(WORD_NS, "w:t");
            t.setAttribute("xml:space", "preserve");
            t.textContent = seg.text;
            run.appendChild(t);
            insWrapper.appendChild(run);
        }

        parent.insertBefore(insWrapper, targetSpan.runElement.nextSibling);
    } catch (e: any) {
        // Silent catch in production
    }
}

// ==========================================
// MODE 2: RECONSTRUCTION MODE (Full Legacy Implementation)
// Restored from Legacy Vibe 3.3 Stable
// ==========================================

function applyReconstructionMode(
    xmlDoc: Document,
    originalText: string,
    modifiedText: string,
    serializer: XMLSerializer,
    author: string
): { oxml: string; hasChanges: boolean } {
    const body = xmlDoc.getElementsByTagName("w:body")[0] || xmlDoc.documentElement;
    const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));

    if (paragraphs.length === 0) return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };

    let originalFullText = "";
    const propertyMap: Array<{ start: number; end: number; rPr: Element | null; wrapper?: Element }> = [];
    const paragraphMap: Array<{ start: number; end: number; pPr: Element | null; container: Element }> = [];
    const sentinelMap: Array<{ start: number; node: Node; isTextBox?: boolean; originalContainer?: Element }> = [];
    const referenceMap = new Map<string, Node>();
    const tokenToCharMap = new Map<string, string>();
    let nextCharCode = 0xe000;

    const uniqueContainers = new Set<Element>();
    const replacementContainers = new Map<Element, Element>();

    paragraphs.forEach((p, pIndex) => {
        const pStart = originalFullText.length;
        const children = Array.from(p.childNodes);

        children.forEach((child) => {
            if (child.nodeName === "w:r") {
                const r = child as Element;
                const rPr = r.getElementsByTagName("w:rPr")[0] || null;
                const runChildren = Array.from(r.childNodes);

                runChildren.forEach((rc) => {
                    if (rc.nodeName === "w:t") {
                        const textContent = (rc as Element).textContent || "";
                        if (textContent.length > 0) {
                            propertyMap.push({
                                start: originalFullText.length,
                                end: originalFullText.length + textContent.length,
                                rPr: rPr,
                            });
                            originalFullText += textContent;
                        }
                    } else if (["w:drawing", "w:pict", "w:object", "w:fldChar", "w:instrText"].includes(rc.nodeName)) {
                        const rcElement = rc as Element;
                        const txbxContent = rcElement.getElementsByTagName("w:txbxContent")[0];
                        const hasTextBox = rc.nodeName === "w:pict" && !!txbxContent;

                        if (hasTextBox) {
                            sentinelMap.push({
                                start: originalFullText.length,
                                node: rc,
                                isTextBox: true,
                                originalContainer: txbxContent as Element,
                            });
                            originalFullText += "\uFFFC";
                            propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
                        } else {
                            sentinelMap.push({ start: originalFullText.length, node: rc });
                            originalFullText += "\uFFFC";
                            propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
                        }
                    } else if (rc.nodeName === "w:footnoteReference" || rc.nodeName === "w:endnoteReference") {
                        const ref = rc as Element;
                        const id = ref.getAttribute("w:id");
                        if (id) {
                            const type = rc.nodeName === "w:footnoteReference" ? "FN" : "EN";
                            const tokenString = `{{__${type}_${id}__}}`;
                            const char = String.fromCharCode(nextCharCode++);
                            referenceMap.set(char, rc);
                            tokenToCharMap.set(tokenString, char);
                            originalFullText += char;
                            propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
                        }
                    }
                });
            } else if (child.nodeName === "w:hyperlink") {
                const h = child as Element;
                const hChildren = Array.from(h.childNodes);
                hChildren.forEach((hc) => {
                    if (hc.nodeName === "w:r") {
                        const r = hc as Element;
                        const rPr = r.getElementsByTagName("w:rPr")[0] || null;
                        const texts = Array.from(r.getElementsByTagName("w:t"));
                        texts.forEach((t) => {
                            const textContent = t.textContent || "";
                            if (textContent.length > 0) {
                                propertyMap.push({
                                    start: originalFullText.length,
                                    end: originalFullText.length + textContent.length,
                                    rPr: rPr,
                                    wrapper: h,
                                });
                                originalFullText += textContent;
                            }
                        });
                    }
                });
            } else if (["w:sdt", "w:oMath", "m:oMath", "w:bookmarkStart", "w:bookmarkEnd"].includes(child.nodeName)) {
                sentinelMap.push({ start: originalFullText.length, node: child });
                originalFullText += "\uFFFC";
            }
        });

        if (pIndex < paragraphs.length - 1) {
            originalFullText += "\n";
        }

        const pEnd = originalFullText.length;
        const pPr = (p as Element).getElementsByTagName("w:pPr")[0] || null;
        const container = p.parentNode as Element;
        if (container) uniqueContainers.add(container);

        paragraphMap.push({
            start: pStart,
            end: pEnd,
            pPr: pPr,
            container: container || body as Element,
        });
    });

    let processedModifiedText = modifiedText || "";

    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, "g"), char);
    });

    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(originalFullText, processedModifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // Post-process to ensure whole-word replacements
    const wordCleanedDiffs = cleanupDiffsToWords(diffs);

    const containerFragments = new Map<Element, DocumentFragment>();
    uniqueContainers.forEach((c) => containerFragments.set(c, xmlDoc.createDocumentFragment()));
    if (!containerFragments.has(body as Element)) containerFragments.set(body as Element, xmlDoc.createDocumentFragment());

    const getParagraphInfo = (index: number): { pPr: Element | null; container: Element } => {
        const match = paragraphMap.find((m) => index >= m.start && index < m.end);
        if (!match && paragraphMap.length > 0) {
            const last = paragraphMap[paragraphMap.length - 1];
            return { pPr: last.pPr, container: last.container };
        }
        return match ? { pPr: match.pPr, container: match.container } : { pPr: null, container: body as Element };
    };

    const createNewParagraph = (pPr: Element | null): Element => {
        const newP = xmlDoc.createElement("w:p");
        if (pPr) newP.appendChild(pPr.cloneNode(true));
        return newP;
    };

    let startInfo = getParagraphInfo(0);
    let currentParagraph = createNewParagraph(startInfo.pPr);
    let currentContainer = startInfo.container;
    let currentFragment = containerFragments.get(currentContainer);
    if (currentFragment) currentFragment.appendChild(currentParagraph);

    let currentOriginalIndex = 0;
    const usedReferenceTokens = new Set<string>();

    const getRunProperties = (index: number): { rPr: Element | null; wrapper?: Element } => {
        const match = propertyMap.find((m) => index >= m.start && index < m.end);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const appendTextToCurrent = (text: string, type: "equal" | "insert" | "delete", rPr: Element | null, wrapper: Element | undefined, baseIndex: number): void => {
        const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);
        let localOffset = 0;

        parts.forEach((part) => {
            if (part === "\n") {
                if (type !== "delete") {
                    let pPr: Element | null = null;
                    let targetContainer = currentContainer;

                    if (type === "equal") {
                        const lookupIndex = baseIndex + localOffset + 1;
                        const info = getParagraphInfo(lookupIndex);
                        pPr = info.pPr;
                        targetContainer = info.container;
                    } else {
                        const info = getParagraphInfo(baseIndex);
                        pPr = info.pPr;
                        targetContainer = info.container;
                    }

                    currentParagraph = createNewParagraph(pPr);
                    if (targetContainer !== currentContainer) {
                        currentContainer = targetContainer;
                        currentFragment = containerFragments.get(currentContainer);
                    }
                    if (currentFragment) currentFragment.appendChild(currentParagraph);
                }
                localOffset += 1;
            } else if (part === "\uFFFC") {
                const sentinelIndex = baseIndex + localOffset;
                const sentinel = sentinelMap.find((s) => s.start === sentinelIndex);

                if (sentinel) {
                    const clone = sentinel.node.cloneNode(true) as Element;
                    if (sentinel.isTextBox && sentinel.originalContainer) {
                        const newContainer = clone.getElementsByTagName("w:txbxContent")[0];
                        if (newContainer) {
                            while (newContainer.firstChild) newContainer.removeChild(newContainer.firstChild);
                            replacementContainers.set(sentinel.originalContainer, newContainer as Element);
                        }
                    }
                    if (sentinel.node.nodeName === "w:r" || sentinel.node.parentNode?.nodeName === "w:r") {
                        const run = xmlDoc.createElement("w:r");
                        if (rPr) run.appendChild(rPr.cloneNode(true));
                        run.appendChild(clone);
                        currentParagraph.appendChild(run);
                    } else {
                        currentParagraph.appendChild(clone);
                    }
                }
                localOffset += 1;
            } else if (referenceMap.has(part)) {
                if (type !== "delete") {
                    const refNode = referenceMap.get(part);
                    if (refNode) {
                        const clone = refNode.cloneNode(true);
                        const run = xmlDoc.createElement("w:r");
                        if (rPr) run.appendChild(rPr.cloneNode(true));
                        run.appendChild(clone);
                        currentParagraph.appendChild(run);
                        let tokenString = "";
                        tokenToCharMap.forEach((c, t) => {
                            if (c === part) tokenString = t;
                        });
                        usedReferenceTokens.add(tokenString);
                    }
                }
                localOffset += part.length;
            } else if (part.length > 0) {
                let parent: Element = currentParagraph;
                if (wrapper) {
                    const wrapperClone = wrapper.cloneNode(false) as Element;
                    parent = wrapperClone;
                    currentParagraph.appendChild(wrapperClone);
                }

                // MARKDOWN PARSING: For insertions, parse **bold** and *italic*
                if (type === "insert") {
                    const segments = parseFormattingMarkers(part);
                    const ins = xmlDoc.createElement("w:ins");
                    ins.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
                    ins.setAttribute("w:author", author);
                    ins.setAttribute("w:date", new Date().toISOString());

                    for (const seg of segments) {
                        const run = xmlDoc.createElement("w:r");

                        // Clone base rPr and add formatting
                        const runRPr = rPr ? rPr.cloneNode(true) as Element : xmlDoc.createElement("w:rPr");
                        if (seg.formatting.bold) {
                            const b = xmlDoc.createElementNS(WORD_NS, "w:b");
                            runRPr.appendChild(b);
                        }
                        if (seg.formatting.italic) {
                            const i = xmlDoc.createElementNS(WORD_NS, "w:i");
                            runRPr.appendChild(i);
                        }
                        if (seg.formatting.underline) {
                            const u = xmlDoc.createElementNS(WORD_NS, "w:u");
                            u.setAttribute("w:val", "single");
                            runRPr.appendChild(u);
                        }
                        if (runRPr.childNodes.length > 0 || rPr) {
                            run.appendChild(runRPr);
                        }

                        const t = xmlDoc.createElement("w:t");
                        t.setAttribute("xml:space", "preserve");
                        t.textContent = seg.text;
                        run.appendChild(t);
                        ins.appendChild(run);
                    }
                    parent.appendChild(ins);
                } else if (type === "delete") {
                    const run = xmlDoc.createElement("w:r");
                    if (rPr) run.appendChild(rPr.cloneNode(true));

                    const t = xmlDoc.createElement("w:delText");
                    t.setAttribute("xml:space", "preserve");
                    t.textContent = part;
                    run.appendChild(t);

                    const del = xmlDoc.createElement("w:del");
                    del.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
                    del.setAttribute("w:author", author);
                    del.setAttribute("w:date", new Date().toISOString());
                    del.appendChild(run);
                    parent.appendChild(del);
                } else {
                    const run = xmlDoc.createElement("w:r");
                    if (rPr) run.appendChild(rPr.cloneNode(true));

                    const t = xmlDoc.createElement("w:t");
                    t.setAttribute("xml:space", "preserve");
                    t.textContent = part;
                    run.appendChild(t);
                    parent.appendChild(run);
                }
                localOffset += part.length;
            }
        });
    };

    for (const diff of wordCleanedDiffs) {
        const [op, text] = diff;
        if (op === 0) {
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const currentPropRange = propertyMap.find(
                    (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
                );
                let length = currentPropRange
                    ? Math.min(currentPropRange.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);
                appendTextToCurrent(chunk, "equal", props.rPr || null, props.wrapper, currentOriginalIndex + offset);
                offset += length;
            }
            currentOriginalIndex += text.length;
        } else if (op === 1) {
            const isStartOfParagraph = paragraphMap.some((p) => p.start === currentOriginalIndex);
            const props =
                currentOriginalIndex > 0 && !isStartOfParagraph
                    ? getRunProperties(currentOriginalIndex - 1)
                    : getRunProperties(currentOriginalIndex);
            appendTextToCurrent(text, "insert", props.rPr || null, props.wrapper, currentOriginalIndex);
        } else if (op === -1) {
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const currentPropRange = propertyMap.find(
                    (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
                );
                let length = currentPropRange
                    ? Math.min(currentPropRange.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);
                appendTextToCurrent(chunk, "delete", props.rPr || null, props.wrapper, currentOriginalIndex + offset);
                offset += length;
            }
            currentOriginalIndex += text.length;
        }
    }

    paragraphs.forEach((p) => {
        if (p.parentNode) p.parentNode.removeChild(p);
    });

    containerFragments.forEach((fragment, container) => {
        const replacement = replacementContainers.get(container);
        const target = replacement || container;
        target.appendChild(fragment);
    });

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

// ==========================================
// MODE 3: PARAGRAPH-AWARE MODE (Multi-paragraph)
// ==========================================

function applyParagraphAwareMode(
    xmlDoc: Document,
    originalText: string,
    modifiedText: string,
    serializer: XMLSerializer,
    author: string
): { oxml: string; hasChanges: boolean } {
    const allParagraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
    if (allParagraphs.length === 0) {
        logWarn("[Engine] No paragraphs found in OXML");
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    // Extract text from each paragraph in the OXML
    const oxmlParaTexts: string[] = [];
    const oxmlParaElements: Element[] = [];

    for (const p of allParagraphs) {
        let paraText = "";
        const runs = p.getElementsByTagName("w:t");
        for (let i = 0; i < runs.length; i++) {
            paraText += runs[i].textContent || "";
        }
        oxmlParaTexts.push(paraText.trim());
        oxmlParaElements.push(p as Element);
    }

    // Split modified text into paragraphs (by newline)
    const modifiedParas = modifiedText.split(/\n/).map(p => p.trim()).filter(p => p.length > 0);

    log("[Engine] Paragraph-aware mode: " + oxmlParaTexts.length + " original, " + modifiedParas.length + " modified");

    // ID ECHO PROTOCOL: PRIMARY MATCHING STRATEGY
    const hasIdTags = /<P\d+>/.test(modifiedText);

    interface ParaMatch {
        origIdx: number;
        modIdx: number;
        score: number;
    }

    const matches: ParaMatch[] = [];
    const usedOrig = new Set<number>();
    const usedMod = new Set<number>();

    if (hasIdTags) {
        log("[Engine] ID Echo Protocol: parsing ID tags for deterministic matching");

        const tagPattern = /<P(\d+)>([\s\S]*?)<\/P\1>/g;
        let tagMatch;
        const modifiedById = new Map<number, { idx: number; content: string }>();

        let modIdx = 0;
        while ((tagMatch = tagPattern.exec(modifiedText)) !== null) {
            const paraId = parseInt(tagMatch[1], 10);
            const content = tagMatch[2].trim();
            modifiedById.set(paraId, { idx: modIdx, content });
            modIdx++;
        }

        for (const [paraId, modInfo] of modifiedById.entries()) {
            const origIdx = paraId - 1; // Convert 1-indexed to 0-indexed

            if (origIdx >= 0 && origIdx < oxmlParaTexts.length) {
                matches.push({ origIdx, modIdx: modInfo.idx, score: 1.0 });
                usedOrig.add(origIdx);
                usedMod.add(modInfo.idx);
            }
        }
    }

    // FALLBACK: Fuzzy similarity matching
    if (!hasIdTags || matches.length === 0) {
        log("[Engine] Using fuzzy similarity matching");

        const matchThreshold = 0.3;

        function normalizeForComparison(text: string): string {
            return text
                .replace(/^\s*\d+(\.\d+)*\.?\s*/g, '')
                .replace(/^\s*\([a-z]\)\s*/gi, '')
                .replace(/^\s*[a-z]\)\s*/gi, '')
                .replace(/\*\*(.+?)\*\*/g, '$1')
                .toLowerCase()
                .trim();
        }

        function similarity(a: string, b: string): number {
            if (!a || !b) return 0;
            const aNorm = normalizeForComparison(a);
            const bNorm = normalizeForComparison(b);
            if (aNorm === bNorm) return 1.0;
            if (!aNorm || !bNorm) return 0;

            const aWords = new Set(aNorm.split(/\s+/).filter(w => w.length > 2));
            const bWords = new Set(bNorm.split(/\s+/).filter(w => w.length > 2));
            if (aWords.size === 0 || bWords.size === 0) return 0;

            let overlap = 0;
            for (const word of aWords) {
                if (bWords.has(word)) overlap++;
            }
            return overlap / Math.max(aWords.size, bWords.size);
        }

        const allScores: ParaMatch[] = [];
        for (let i = 0; i < oxmlParaTexts.length; i++) {
            if (usedOrig.has(i)) continue;
            for (let j = 0; j < modifiedParas.length; j++) {
                if (usedMod.has(j)) continue;
                const score = similarity(oxmlParaTexts[i], modifiedParas[j]);
                if (score >= matchThreshold) {
                    allScores.push({ origIdx: i, modIdx: j, score });
                }
            }
        }

        allScores.sort((a, b) => {
            if (b.score !== a.score) return b.score - a.score;
            return (a.origIdx + a.modIdx) - (b.origIdx + b.modIdx);
        });

        for (const m of allScores) {
            if (!usedOrig.has(m.origIdx) && !usedMod.has(m.modIdx)) {
                matches.push(m);
                usedOrig.add(m.origIdx);
                usedMod.add(m.modIdx);
            }
        }
    }

    // Process each original paragraph
    let hasChanges = false;

    for (let i = 0; i < oxmlParaElements.length; i++) {
        const p = oxmlParaElements[i];
        const match = matches.find(m => m.origIdx === i);

        if (!match) {
            // DELETED
            log("[Engine] Deleting paragraph " + i);
            wrapParagraphContentInDel(xmlDoc, p, author);
            hasChanges = true;
        } else {
            // MATCHED
            const modText = modifiedParas[match.modIdx];
            const origText = oxmlParaTexts[i];

            if (modText.trim() !== origText.trim()) {
                log("[Engine] Modifying paragraph " + i);
                applyCharacterDiffToParagraph(xmlDoc, p, origText, modText, author);
                hasChanges = true;
            }
        }
    }

    // Handle NEW paragraphs
    const newParaIndices: number[] = [];
    for (let j = 0; j < modifiedParas.length; j++) {
        if (!usedMod.has(j)) {
            newParaIndices.push(j);
        }
    }

    if (newParaIndices.length > 0) {
        log("[Engine] " + newParaIndices.length + " new paragraphs to insert");
        const lastPara = oxmlParaElements[oxmlParaElements.length - 1];
        const lastParaParent = lastPara.parentNode;

        let refRPr: Element | null = null;
        const lastRuns = lastPara.getElementsByTagName("w:r");
        if (lastRuns.length > 0) {
            refRPr = lastRuns[0].getElementsByTagName("w:rPr")[0] as Element || null;
        }

        for (const idx of newParaIndices) {
            const newText = modifiedParas[idx];
            const newP = createInsertedParagraph(xmlDoc, newText, refRPr, author);
            if (lastParaParent && lastPara.nextSibling) {
                lastParaParent.insertBefore(newP, lastPara.nextSibling);
            } else if (lastParaParent) {
                lastParaParent.appendChild(newP);
            }
            hasChanges = true;
        }
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges };
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

export function sanitizeAiResponse(text: string): string {
    let cleaned = text;
    cleaned = cleaned.replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, "");
    cleaned = cleaned.replace(/\$\\text\{/g, "").replace(/\}\$/g, "");
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, "$1");
    return cleaned;
}

export function sanitizeForNumberedList(originalText: string, aiResponse: string, wasNumberedList: boolean): string {
    let result = sanitizeAiResponse(aiResponse);

    if (wasNumberedList) {
        const originalHasNumber = /^\s*(\d+[.\)]|\([a-z\d]\)|[a-z][.\)])/i.test(originalText);

        if (!originalHasNumber) {
            result = stripAiAddedNumbering(result);
        }
    }

    return result;
}

/**
 * Restore manual clause number if AI stripped it.
 * For paragraphs WITHOUT auto-numbering (no w:numPr), the clause number
 * is part of the text content. If AI returns text without the number,
 * we prepend it back.
 */
export function restoreManualNumber(originalText: string, aiText: string, isListItem: boolean): string {
    if (isListItem) return aiText;

    const patterns = [
        /^(\((?:\d+\.)+\d*\))\s*/,     // "(3.1.2) "
        /^(\(\d+\))\s*/,               // "(3) "
        /^((?:\d+\.)+\d*\.?)\s+/,      // "3.1.2 " or "3.1.2. "
        /^(\d+\.)\s+/,                 // "3. "
        /^\(([a-z]|[ivx]+)\)\s+/i,     // "(a) " or "(i) "
        /^([a-z]|[ivx]+)\)\s+/i,       // "a) " or "i) "
    ];

    for (const pattern of patterns) {
        const m = originalText.match(pattern);
        if (m && !aiText.startsWith(m[0]) && !aiText.match(pattern)) {
            log(" Restoring manual clause number: " + m[0].trim());
            return m[0] + aiText;
        }
    }
    return aiText;
}

function stripAiAddedNumbering(text: string): string {
    const patterns = [
        /^\s*\d+[.\)]\s+/,
        /^\s*\(\d+\)\s+/,
        /^\s*\([a-z]\)\s+/i,
        /^\s*[a-z][.\)]\s+/i,
        /^\s*\([ivxlcdm]+\)\s+/i,
        /^\s*[ivxlcdm]+[.\)]\s+/i,
        /^\s*[-*]\s+/,
    ];

    for (const pattern of patterns) {
        if (pattern.test(text)) {
            return text.replace(pattern, '');
        }
    }

    return text;
}

function hasNumberingInOxml(xmlDoc: Document): boolean {
    if (xmlDoc.getElementsByTagName("w:numPr").length > 0) return true;

    try {
        if (xmlDoc.getElementsByTagNameNS(WORD_NS, "numPr").length > 0) return true;
    } catch (e) { /* Ignore */ }

    const pPrs = xmlDoc.getElementsByTagName("w:pPr");
    for (let i = 0; i < pPrs.length; i++) {
        const children = pPrs[i].childNodes;
        for (let j = 0; j < children.length; j++) {
            const nodeName = (children[j] as Element).nodeName || "";
            if (nodeName === "w:numPr" || nodeName.endsWith(":numPr")) {
                return true;
            }
        }
    }

    return false;
}

function cleanupDiffsToWords(diffs: [number, string][]): [number, string][] {
    if (diffs.length < 2) return diffs;

    const isWordChar = (c: string) => /[a-zA-Z0-9]/.test(c);
    const workingDiffs: [number, string][] = diffs.map(d => [d[0], d[1]]);
    const result: [number, string][] = [];

    for (let i = 0; i < workingDiffs.length; i++) {
        const [op, text] = workingDiffs[i];

        if (op === 0) {
            let equalText = text;

            if (i + 1 < workingDiffs.length && workingDiffs[i + 1][0] === -1) {
                const deleteText = workingDiffs[i + 1][1];

                if (equalText.length > 0 && isWordChar(equalText.slice(-1)) &&
                    deleteText.length > 0 && isWordChar(deleteText[0])) {

                    let wordStart = equalText.length;
                    while (wordStart > 0 && isWordChar(equalText[wordStart - 1])) {
                        wordStart--;
                    }
                    const sharedPrefix = equalText.substring(wordStart);

                    if (sharedPrefix.length > 0) {
                        workingDiffs[i + 1] = [-1, sharedPrefix + deleteText];

                        if (i + 2 < workingDiffs.length && workingDiffs[i + 2][0] === 1) {
                            const insertText = workingDiffs[i + 2][1];
                            if (insertText.length > 0 && isWordChar(insertText[0])) {
                                workingDiffs[i + 2] = [1, sharedPrefix + insertText];
                            }
                        } else {
                            workingDiffs.splice(i + 2, 0, [1, sharedPrefix]);
                        }

                        equalText = equalText.substring(0, wordStart);
                    }
                }
            }

            if (result.length > 0 && equalText.length > 0) {
                const lastResult = result[result.length - 1];

                if (lastResult[0] === 1 && lastResult[1].length > 0 &&
                    isWordChar(lastResult[1].slice(-1)) && isWordChar(equalText[0])) {

                    let wordEnd = 0;
                    while (wordEnd < equalText.length && isWordChar(equalText[wordEnd])) {
                        wordEnd++;
                    }
                    const sharedSuffix = equalText.substring(0, wordEnd);

                    if (sharedSuffix.length > 0) {
                        result[result.length - 1] = [1, lastResult[1] + sharedSuffix];

                        if (result.length >= 2 && result[result.length - 2][0] === -1) {
                            const prevDelete = result[result.length - 2][1];
                            if (prevDelete.length > 0 && isWordChar(prevDelete.slice(-1))) {
                                result[result.length - 2] = [-1, prevDelete + sharedSuffix];
                            }
                        }

                        equalText = equalText.substring(wordEnd);
                    }
                }
            }

            if (equalText.length > 0) {
                result.push([0, equalText]);
            }
        } else {
            result.push([op, text]);
        }
    }

    const merged: [number, string][] = [];
    for (const [op, text] of result) {
        if (text.length === 0) continue;

        if (merged.length > 0 && merged[merged.length - 1][0] === op) {
            merged[merged.length - 1] = [op, merged[merged.length - 1][1] + text];
        } else {
            merged.push([op, text]);
        }
    }

    return merged;
}

function wrapParagraphContentInDel(xmlDoc: Document, p: Element, author: string): void {
    const runs = Array.from(p.getElementsByTagName("w:r"));

    for (const r of runs) {
        if (r.parentNode?.nodeName === "w:del") continue;

        const parent = r.parentNode;
        if (!parent) continue;

        const textNodes = r.getElementsByTagName("w:t");
        for (let i = 0; i < textNodes.length; i++) {
            const t = textNodes[i];
            const delText = xmlDoc.createElementNS(WORD_NS, "w:delText");
            delText.textContent = t.textContent;
            delText.setAttribute("xml:space", "preserve");
            t.parentNode?.replaceChild(delText, t);
        }

        const delWrapper = xmlDoc.createElementNS(WORD_NS, "w:del");
        delWrapper.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
        delWrapper.setAttribute("w:author", author);
        delWrapper.setAttribute("w:date", new Date().toISOString());

        parent.insertBefore(delWrapper, r);
        delWrapper.appendChild(r);
    }
}

function applyCharacterDiffToParagraph(xmlDoc: Document, p: Element, origText: string, modText: string, author: string): void {
    const runs = Array.from(p.getElementsByTagName("w:r"));
    if (runs.length === 0) return;

    const refRPr = runs[0].getElementsByTagName("w:rPr")[0] as Element || null;
    const firstRun = runs[0];
    const parent = firstRun.parentNode;
    if (!parent) return;

    for (const r of runs) {
        if (r.parentNode) r.parentNode.removeChild(r);
    }

    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(origText, modText);
    dmp.diff_cleanupSemantic(diffs);
    const cleanedDiffs = cleanupDiffsToWords(diffs);

    for (const [op, text] of cleanedDiffs) {
        if (text.length === 0) continue;

        if (op === 0) {
            const run = createTextRun(xmlDoc, text, refRPr, false);
            parent.appendChild(run);
        } else if (op === -1) {
            const delRun = createTextRun(xmlDoc, text, refRPr, true);
            const delWrapper = createTrackChange(xmlDoc, "del", delRun, author);
            parent.appendChild(delWrapper);
        } else if (op === 1) {
            const insRun = createTextRun(xmlDoc, text, refRPr, false);
            const insWrapper = createTrackChange(xmlDoc, "ins", insRun, author);
            parent.appendChild(insWrapper);
        }
    }
}

function createInsertedParagraph(xmlDoc: Document, text: string, refRPr: Element | null, author: string): Element {
    const p = xmlDoc.createElement("w:p");

    const r = xmlDoc.createElement("w:r");
    if (refRPr) {
        r.appendChild(refRPr.cloneNode(true));
    }
    const t = xmlDoc.createElement("w:t");
    t.textContent = text;
    t.setAttribute("xml:space", "preserve");
    r.appendChild(t);

    const insWrapper = xmlDoc.createElement("w:ins");
    insWrapper.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
    insWrapper.setAttribute("w:author", author);
    insWrapper.setAttribute("w:date", new Date().toISOString());
    insWrapper.appendChild(r);

    p.appendChild(insWrapper);
    return p;
}

function createTextRun(xmlDoc: Document, text: string, rPr: Element | null, isDeleted: boolean): Element {
    const run = xmlDoc.createElementNS(WORD_NS, "w:r");
    if (rPr) {
        run.appendChild(rPr.cloneNode(true));
    }
    const textElement = xmlDoc.createElementNS(WORD_NS, isDeleted ? "w:delText" : "w:t");
    textElement.setAttribute("xml:space", "preserve");
    textElement.textContent = text;
    run.appendChild(textElement);
    return run;
}

function createTrackChange(xmlDoc: Document, type: "ins" | "del", runElement: Element | null, author: string): Element {
    const wrapper = xmlDoc.createElementNS(WORD_NS, "w:" + type);
    wrapper.setAttribute("w:id", Math.floor(Math.random() * 10000000).toString());
    wrapper.setAttribute("w:author", author);
    wrapper.setAttribute("w:date", new Date().toISOString());
    if (runElement) {
        wrapper.appendChild(runElement);
    }
    return wrapper;
}
