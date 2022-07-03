import { XElement } from "@openxmldev/linq-to-xml";
import { W } from "@openxmldev/linq-to-ooxml";
import { getDocument } from "./flat-opc";

/* global Word */

/**
 * Represents a selection.
 */
export interface Selection {
  /**
   * The range representing the user's selection, which will be extended as necessary
   * to cover full paragraphs.
   */
  range: Word.Range;

  /**
   * The pkg:package root element of the Flat OPC package representing the contents
   * of the selected range.
   */
  packageRootElement: XElement;

  /**
   * The w:document root element of the main document part.
   */
  document: XElement;

  /**
   * The flag indicating whether the selection is valid for the purpose of replacing
   * the contents of the selected range with the Flat OPC package in its original
   * or transformed form.
   */
  isValid: boolean;
}

/**
 * Transforms the user-selected range, which is extended to cover one or more complete
 * paragraphs.
 *
 * @param context The Word context.
 * @returns true, if the extended selection had its formatting reset; false, otherwise.
 */
export async function transformSelection(
  transform: (selection: Selection) => void,
  context: Word.RequestContext
): Promise<boolean> {
  // Step 1: Get a valid selection, i.e., one that can be transformed without messing
  // up the document.
  const selection = await getSelection(context);
  if (!selection.isValid) return false;

  // Step 2: Transform the Flat OPC package representing the contents of the selection.
  transform(selection);

  // Step 3: Insert the transformed Flat OPC package into the selected range,
  // replacing the contents of the selection.
  await insertOoxml(selection, context);

  return true;
}

/**
 * Gets a Selection object, including the selected range, the Flat OPC package representing
 * the contents of the selected range, the w:document root element of the document part,
 * and a flag indicating whether that selection is valid for the purpose of replacing the
 * contents of the selected range with the Flat OPC package in its original or transformed
 * form.
 *
 * @param context The Word context.
 * @returns A Promise<Selection>.
 */
async function getSelection(context: Word.RequestContext): Promise<Selection> {
  const range: Word.Range = selectCompleteParagraphs(context);
  const ooxmlResult = range.getOoxml();
  await context.sync();

  // eslint-disable-next-line office-addins/load-object-before-read
  const packageRootElement: XElement = XElement.parse(ooxmlResult.value);
  const document: XElement = getDocument(packageRootElement);
  const isValid = selectionIsValid(document);

  return { range, packageRootElement, document, isValid };
}

/**
 * Using the given selection, replaces the contents of the range with the Flat OPC package.
 *
 * @param selection The (partial) Selection object containing the range and Flat OPC package.
 * @param context The Word context.
 */
async function insertOoxml(selection: Selection, context: Word.RequestContext): Promise<void> {
  const ooxml = selection.packageRootElement.toString();
  selection.range.insertOoxml(ooxml, Word.InsertLocation.replace);
  await context.sync();
}

/**
 * Based on the user's selection, extends the selection to cover complete paragraphs.
 *
 * @param context The Word context.
 * @returns A range consisting of complete paragraphs.
 */
function selectCompleteParagraphs(context: Word.RequestContext): Word.Range {
  const selection = context.document.getSelection();
  const firstParagraph = selection.paragraphs.getFirst();
  const lastParagraph = selection.paragraphs.getLast();

  // We need to use 'Whole' in order to select the paragraph(s) including the
  // paragraph marks. If we don't select the paragraph mark, we will not get
  // the w:pPr element!
  const selectionStart = firstParagraph.getRange("Whole");
  const selectionEnd = lastParagraph.getRange("Whole");

  return selection.expandTo(selectionStart).expandTo(selectionEnd);
}

/**
 * Determines whether the given w:document element represents a valid selection that
 * can later be replaced without messing up the document.
 *
 * @param document The main document part.
 * @returns true, if the selection is valid; false, otherwise.
 */
function selectionIsValid(document: XElement): boolean {
  // The w:body element will always have at least three child elements:
  // - one or more selected elements (e.g., w:p, w:tbl) and
  // - the following two final elements:
  //   - a w:p element with an attribute w:rsidR="00000000", and
  //   - a w:sectPr element with an attribute w:rsidR="00000000".
  // If the last selected element (the one that precedes the final w:p and w:sectPr elements)
  // is a w:tbl, replacing the OOXML of the selected range will mess up the document.
  var lastSelectedElement = document.elements(W.body).elements().takeLast(3).first();
  return lastSelectedElement.name !== W.tbl;
}
