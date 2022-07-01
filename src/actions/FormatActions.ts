/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable office-addins/call-sync-before-read */
/* eslint-disable no-undef */
import { XNode, XElement } from "@openxmldev/linq-to-xml";
import { W } from "@openxmldev/linq-to-ooxml";
import { PKG } from "./PKG";

const mainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

export class FormatActions {
  constructor(private context: Word.RequestContext) {}

  async reset(): Promise<void> {
    let selection: Word.Range = await this.selectCompleteParagraphs();
    const ooxmlResult = selection.getOoxml();

    await this.context.sync();

    const root: XElement = XElement.parse(ooxmlResult.value);
    const part: XElement = root.elements().first((e) => e.attribute(PKG.contentType)?.value === mainContentType);

    const transformedPart: XElement = removeDirectFormatting(part);

    // Important note: This simplified example does not consider tables, which require special treatment.
    // transforming a paragraph within a table cell, for example, will not lead to the desired result
    // without further steps (e.g., removing the table in which said paragraph will be wrapped).
    part.replaceWith(transformedPart);

    const ooxml = root.toString();
    selection = selection.insertOoxml(ooxml, Word.InsertLocation.replace);
    await this.context.sync();
  }

  private async selectCompleteParagraphs(): Promise<Word.Range> {
    const selection = this.context.document.getSelection();
    const firstParagraph = selection.paragraphs.getFirst();
    const lastParagraph = selection.paragraphs.getLast();

    // We need to use 'Whole' in order to select the paragraph(s) including the
    // paragraph marks. If we don't select the paragraph mark, we will not get
    // the w:pPr element!
    const selectionStart = firstParagraph.getRange("Whole");
    const selectionEnd = lastParagraph.getRange("Whole");

    return Promise.resolve(selection.expandTo(selectionStart).expandTo(selectionEnd));
  }
}

//
// Pure Functional Transformations
//

/**
 * Removes direct paragraph and run formatting.
 *
 * @param element The element to be transformed.
 * @returns A new transformed element.
 */
function removeDirectFormatting(element: XElement): XElement {
  return removeDirectFormattingTransformation(element) as XElement;
}

function removeDirectFormattingTransformation(node: XNode): XNode | null {
  if (!(node instanceof XElement)) return node;

  const element: XElement = node;

  if (element.name === W.pPr) {
    return paragraphPropertiesTransformation(element);
  }

  if (element.name === W.rPr) {
    return runPropertiesTransformation(element);
  }

  // Perform identity transformation.
  return new XElement(element.name, element.attributes(), element.nodes().select(removeDirectFormattingTransformation));
}

function paragraphPropertiesTransformation(element: XElement): XElement | null {
  if (element.name === W.pPr) {
    const retainedElements: XElement[] = element
      .elements()
      .select(paragraphPropertiesTransformation)
      .where((e) => e !== null)
      .toArray();

    return retainedElements.length > 0 ? new XElement(W.pPr, retainedElements) : null;
  }

  if (element.name === W.rPr) {
    return runPropertiesTransformation(element);
  }

  if (element.name === W.pStyle || element.name === W.numPr) {
    return element;
  }

  return null;
}

function runPropertiesTransformation(element: XElement): XElement | null {
  if (element.name === W.rPr) {
    const retainedElements: XElement[] = element
      .elements()
      .select(runPropertiesTransformation)
      .where((e) => e !== null)
      .toArray();

    return retainedElements.length > 0 ? new XElement(W.rPr, retainedElements) : null;
  }

  if (element.name === W.rStyle) {
    return element;
  }

  return null;
}
