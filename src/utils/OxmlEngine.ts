import { diff_match_patch } from 'diff-match-patch';

/**
 * OxmlEngine V5.1 - Hybrid Mode
 * 
 * TWO MODES:
 * - SURGICAL MODE (for tables): Modifies existing elements in place, never creates/deletes structure
 * - RECONSTRUCTION MODE (for body without tables): Allows new paragraphs for list splitting
 * 
 * This hybrid approach gives us the best of both worlds:
 * - Tables are never broken
 * - Lists can still be split
 */

const isNode = typeof process !== 'undefined' && process.versions != null && process.versions.node != null;

interface TextSpan {
    charStart: number;
    charEnd: number;
    textElement: Element;
    runElement: Element;
    paragraph: Element;
    container: Element;
    rPr: Element | null;
}

interface StructureEntry {
    charStart: number;
    charEnd: number;
    paragraph: Element;
    pPr: Element | null;
    container: Element;
    numPr: Element | null;
}

export function applyRedlineToOxml(
    oxml: string,
    originalText: string,
    modifiedText: string
): { oxml: string; hasChanges: boolean } {

    let DOMParserImpl: typeof DOMParser;
    let XMLSerializerImpl: typeof XMLSerializer;

    if (isNode) {
        DOMParserImpl = (global as any).DOMParser;
        XMLSerializerImpl = (global as any).XMLSerializer;
    } else {
        DOMParserImpl = window.DOMParser;
        XMLSerializerImpl = window.XMLSerializer;
    }

    if (!DOMParserImpl || !XMLSerializerImpl) {
        throw new Error("DOMParser or XMLSerializer not available");
    }

    const parser = new DOMParserImpl();
    let xmlDoc: Document;

    try {
        xmlDoc = parser.parseFromString(oxml, "text/xml");
    } catch (e) {
        console.error("[Engine] Failed to parse OXML:", e);
        return { oxml, hasChanges: false };
    }

    const cleanModifiedText = sanitizeAiResponse(modifiedText);

    if (cleanModifiedText.trim() === originalText.trim()) {
        return { oxml, hasChanges: false };
    }

    // Check for tables
    const tables = xmlDoc.getElementsByTagName("w:tbl");
    const hasTables = tables.length > 0;

    if (hasTables) {
        // SURGICAL MODE for tables
        return applySurgicalMode(xmlDoc, originalText, cleanModifiedText, new XMLSerializerImpl());
    } else {
        // RECONSTRUCTION MODE for body without tables
        return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, new XMLSerializerImpl());
    }
}

// ==========================================
// SURGICAL MODE (for tables)
// ==========================================
function applySurgicalMode(
    xmlDoc: Document,
    originalText: string,
    modifiedText: string,
    serializer: XMLSerializer
): { oxml: string; hasChanges: boolean } {

    let fullText = "";
    const textSpans: TextSpan[] = [];

    const allParagraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));

    allParagraphs.forEach((p, pIndex) => {
        const container = p.parentNode as Element;

        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName === "w:r") {
                const r = child as Element;
                const rPr = r.getElementsByTagName("w:rPr")[0] || null;

                Array.from(r.childNodes).forEach(rc => {
                    if (rc.nodeName === "w:t") {
                        const t = rc as Element;
                        const text = t.textContent || "";
                        if (text.length > 0) {
                            textSpans.push({
                                charStart: fullText.length,
                                charEnd: fullText.length + text.length,
                                textElement: t,
                                runElement: r,
                                paragraph: p,
                                container,
                                rPr
                            });
                            fullText += text;
                        }
                    }
                });
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
                                    textSpans.push({
                                        charStart: fullText.length,
                                        charEnd: fullText.length + text.length,
                                        textElement: t,
                                        runElement: r,
                                        paragraph: p,
                                        container,
                                        rPr
                                    });
                                    fullText += text;
                                }
                            }
                        });
                    }
                });
            }
        });

        if (pIndex < allParagraphs.length - 1) {
            fullText += "\n";
        }
    });

    // Compute diff
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(fullText, modifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // Process deletions and insertions
    let currentPos = 0;
    const processedSpans = new Set<Element>();

    for (const [op, text] of diffs) {
        if (op === 0) {
            currentPos += text.length;
        } else if (op === -1) {
            // DELETE
            processDelete(xmlDoc, textSpans, currentPos, currentPos + text.length, processedSpans);
            currentPos += text.length;
        } else if (op === 1) {
            // INSERT
            const textWithoutNewlines = text.replace(/\n/g, ' ');
            if (textWithoutNewlines.trim().length > 0) {
                processInsert(xmlDoc, textSpans, currentPos, textWithoutNewlines, processedSpans);
            }
        }
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

function processDelete(
    xmlDoc: Document,
    textSpans: TextSpan[],
    startPos: number,
    endPos: number,
    processedSpans: Set<Element>
) {
    const affectedSpans = textSpans.filter(s =>
        s.charEnd > startPos && s.charStart < endPos
    );

    for (const span of affectedSpans) {
        if (processedSpans.has(span.textElement)) continue;

        const deleteStart = Math.max(0, startPos - span.charStart);
        const deleteEnd = Math.min(span.charEnd - span.charStart, endPos - span.charStart);

        const originalText = span.textElement.textContent || "";
        const beforeText = originalText.substring(0, deleteStart);
        const deletedText = originalText.substring(deleteStart, deleteEnd);
        const afterText = originalText.substring(deleteEnd);

        if (deletedText.length === 0) continue;

        const parent = span.runElement.parentNode;
        if (!parent) continue;

        if (beforeText.length === 0 && afterText.length === 0) {
            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun);
            parent.insertBefore(delWrapper, span.runElement);
            parent.removeChild(span.runElement);
        } else {
            if (beforeText.length > 0) {
                const beforeRun = createTextRun(xmlDoc, beforeText, span.rPr, false);
                parent.insertBefore(beforeRun, span.runElement);
            }

            const delRun = createTextRun(xmlDoc, deletedText, span.rPr, true);
            const delWrapper = createTrackChange(xmlDoc, 'del', delRun);
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

function processInsert(
    xmlDoc: Document,
    textSpans: TextSpan[],
    pos: number,
    text: string,
    processedSpans: Set<Element>
) {
    // Try to find the span at the insertion position
    let targetSpan = textSpans.find(s => pos >= s.charStart && pos < s.charEnd);

    // If not found, try the span that ends at this position
    if (!targetSpan && pos > 0) {
        targetSpan = textSpans.find(s => pos === s.charEnd);
    }

    // If still not found, try the span just before
    if (!targetSpan && pos > 0) {
        targetSpan = textSpans.find(s => s.charEnd <= pos);
        if (!targetSpan && textSpans.length > 0) {
            // Find the closest span before this position
            const before = textSpans.filter(s => s.charEnd <= pos);
            if (before.length > 0) {
                targetSpan = before[before.length - 1];
            }
        }
    }

    // Last resort: use the last span
    if (!targetSpan && textSpans.length > 0) {
        targetSpan = textSpans[textSpans.length - 1];
    }

    if (targetSpan) {
        // CRITICAL: Ensure we inherit the rPr
        const rPr = targetSpan.rPr;
        const insRun = createTextRun(xmlDoc, text, rPr, false);
        const insWrapper = createTrackChange(xmlDoc, 'ins', insRun);

        const parent = targetSpan.runElement.parentNode;
        if (parent) {
            parent.insertBefore(insWrapper, targetSpan.runElement.nextSibling);
        }
    }
}

// ==========================================
// RECONSTRUCTION MODE (Adapted from Legacy Vibe 3.3 Stable)
function applyReconstructionMode(xmlDoc: Document, originalText: string, modifiedText: string, serializer: XMLSerializer): { oxml: string; hasChanges: boolean } {
    const body = xmlDoc.getElementsByTagName("w:body")[0] || xmlDoc.documentElement;
    const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));

    if (paragraphs.length === 0) return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };

    // 2. BUILD CONTEXT & MAPS
    let originalFullText = "";
    const propertyMap: { start: number; end: number; rPr: Element | null; wrapper?: Node }[] = [];
    const paragraphMap: { start: number; end: number; pPr: Element | null; container: Node }[] = [];
    const sentinelMap: { start: number; node: Node; isTextBox?: boolean; originalContainer?: Node }[] = [];
    const referenceMap = new Map<string, Node>();
    const tokenToCharMap = new Map<string, string>();
    let nextCharCode = 0xe000;

    const uniqueContainers = new Set<Node>();
    const replacementContainers = new Map<Node, Node>();

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
                        const textContent = rc.textContent || "";
                        if (textContent.length > 0) {
                            propertyMap.push({
                                start: originalFullText.length,
                                end: originalFullText.length + textContent.length,
                                rPr: rPr,
                            });
                            originalFullText += textContent;
                        }
                    } else if (["w:drawing", "w:pict", "w:object", "w:fldChar", "w:instrText"].includes(rc.nodeName)) {
                        // SENTINEL LOGIC
                        const rcElement = rc as Element;
                        const txbxContent = rcElement.getElementsByTagName("w:txbxContent")[0];
                        const hasTextBox = rc.nodeName === "w:pict" && !!txbxContent;

                        if (hasTextBox) {
                            sentinelMap.push({
                                start: originalFullText.length,
                                node: rc,
                                isTextBox: true,
                                originalContainer: txbxContent,
                            });
                            originalFullText += "\uFFFC";
                            propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
                        } else {
                            sentinelMap.push({ start: originalFullText.length, node: rc });
                            originalFullText += "\uFFFC";
                            propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
                        }
                    } else if (rc.nodeName === "w:footnoteReference" || rc.nodeName === "w:endnoteReference") {
                        // REFERENCE LOGIC
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
        const pPr = p.getElementsByTagName("w:pPr")[0] || null;
        const container = p.parentNode;
        if (container) uniqueContainers.add(container!);

        paragraphMap.push({
            start: pStart,
            end: pEnd,
            pPr: pPr,
            container: container || body,
        });
    });

    // 3. SANITIZE & PROCESS INPUT
    let processedModifiedText = modifiedText;

    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, "g"), char);
    });

    // 4. DIFF CALCULATION
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(originalFullText, processedModifiedText);
    dmp.diff_cleanupSemantic(diffs);

    // 5. RECONSTRUCTION
    const containerFragments = new Map<Node, DocumentFragment>();
    uniqueContainers.forEach((c) => containerFragments.set(c, xmlDoc.createDocumentFragment()));
    if (!containerFragments.has(body)) containerFragments.set(body, xmlDoc.createDocumentFragment());

    const getParagraphInfo = (index: number): { pPr: Element | null; container: Node } => {
        const match = paragraphMap.find((m) => index >= m.start && index < m.end);
        if (!match && paragraphMap.length > 0) {
            const last = paragraphMap[paragraphMap.length - 1];
            return { pPr: last.pPr, container: last.container };
        }
        return match ? { pPr: match.pPr, container: match.container } : { pPr: null, container: body };
    };

    const createNewParagraph = (pPr: Element | null) => {
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

    const getRunProperties = (index: number): { rPr: Element | null; wrapper?: Node } => {
        const match = propertyMap.find((m) => index >= m.start && index < m.end);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const appendTextToCurrent = (text: string, type: "equal" | "insert" | "delete", rPr: Element | null, wrapper: Node | undefined, baseIndex: number) => {
        const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);
        let localOffset = 0;

        parts.forEach((part) => {
            if (part === "\n") {
                if (type !== "delete") {
                    let pPr: Element | null = null;
                    let targetContainer: Node = currentContainer;

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
                            replacementContainers.set(sentinel.originalContainer, newContainer);
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
                const run = xmlDoc.createElement("w:r");
                if (rPr) run.appendChild(rPr.cloneNode(true));

                let parent: Node = currentParagraph;
                if (wrapper) {
                    const wrapperClone = wrapper.cloneNode(false);
                    parent = wrapperClone;
                    currentParagraph.appendChild(wrapperClone);
                }

                const t = type === "delete" ? xmlDoc.createElement("w:delText") : xmlDoc.createElement("w:t");
                t.setAttribute("xml:space", "preserve");
                t.textContent = part;

                run.appendChild(t);

                if (type === "delete") {
                    const del = xmlDoc.createElement("w:del");
                    del.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
                    del.setAttribute("w:author", "Vibe");
                    del.setAttribute("w:date", new Date().toISOString());
                    del.appendChild(run);
                    parent.appendChild(del);
                } else if (type === "insert") {
                    const ins = xmlDoc.createElement("w:ins");
                    ins.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
                    ins.setAttribute("w:author", "Vibe");
                    ins.setAttribute("w:date", new Date().toISOString());
                    ins.appendChild(run);
                    parent.appendChild(ins);
                } else {
                    parent.appendChild(run);
                }
                localOffset += part.length;
            }
        });
    };

    for (const diff of diffs) {
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
// HELPERS
// ==========================================

function createTrackChange(xmlDoc: Document, type: 'ins' | 'del', run: Element): Element {
    const wrapper = xmlDoc.createElement(type === 'ins' ? "w:ins" : "w:del");
    wrapper.setAttribute("w:id", Math.floor(Math.random() * 90000 + 10000).toString());
    wrapper.setAttribute("w:author", "Vibe Legal");
    wrapper.setAttribute("w:date", new Date().toISOString());
    wrapper.appendChild(run);
    return wrapper;
}

function createTextRun(xmlDoc: Document, text: string, rPr: Element | null, isDelete: boolean): Element {
    const run = xmlDoc.createElement("w:r");
    if (rPr) run.appendChild(rPr.cloneNode(true));

    const textEl = xmlDoc.createElement(isDelete ? "w:delText" : "w:t");
    textEl.setAttribute("xml:space", "preserve");
    textEl.textContent = text;
    run.appendChild(textEl);

    return run;
}

function sanitizeAiResponse(text: string): string {
    let cleaned = text;
    cleaned = cleaned.replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, "");
    cleaned = cleaned.replace(/\$\\text\{/g, "").replace(/\}\$/g, "");
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, "$1");
    return cleaned;
}
