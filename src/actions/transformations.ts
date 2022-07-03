import { XNode, XElement } from "@openxmldev/linq-to-xml";
import { W } from "@openxmldev/linq-to-ooxml";
import { Selection } from "./selections";

/**
 * Removes direct paragraph and run formatting.
 *
 * @param selection A Selection object representing the range to be transformed.
 */
export function removeDirectFormatting(selection: Selection): void {
  const { document } = selection;
  const transformedDocument = removeDirectFormattingTransformation(document);
  document.replaceWith(transformedDocument);
}

//
// Pure Functional Transformations
//

function removeDirectFormattingTransformation(node: XNode): XNode | null {
  // Retain (text) nodes.
  if (!(node instanceof XElement)) return node;

  const element: XElement = node;

  // Transform w:pPr elements.
  if (element.name === W.pPr) {
    return paragraphPropertiesTransformation(element);
  }

  // Transform w:rPr elements.
  if (element.name === W.rPr) {
    return runPropertiesTransformation(element);
  }

  // Perform identity transformation on all other elements.
  return new XElement(element.name, element.attributes(), element.nodes().select(removeDirectFormattingTransformation));
}

function paragraphPropertiesTransformation(element: XElement): XElement | null {
  // Transform w:pPr elements, removing them as well if all children are removed.
  if (element.name === W.pPr) {
    const retainedElements: XElement[] = element
      .elements()
      .select(paragraphPropertiesTransformation)
      .where((e) => e !== null)
      .toArray();

    return retainedElements.length > 0 ? new XElement(W.pPr, retainedElements) : null;
  }

  // Transform w:rPr elements.
  if (element.name === W.rPr) {
    return runPropertiesTransformation(element);
  }

  // Retain w:pStyle and w:numPr elements.
  if (element.name === W.pStyle || element.name === W.numPr) {
    return element;
  }

  // Remove all other elements.
  return null;
}

function runPropertiesTransformation(element: XElement): XElement | null {
  // Transform w:rPr elements, removing them as well if all children are removed.
  if (element.name === W.rPr) {
    const retainedElements: XElement[] = element
      .elements()
      .select(runPropertiesTransformation)
      .where((e) => e !== null)
      .toArray();

    return retainedElements.length > 0 ? new XElement(W.rPr, retainedElements) : null;
  }

  // Retain w:rStyle elements.
  if (element.name === W.rStyle) {
    return element;
  }

  // Remove all other elements.
  return null;
}
