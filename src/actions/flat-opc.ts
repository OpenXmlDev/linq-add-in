import { XElement } from "@openxmldev/linq-to-xml";
import { PKG, W } from "@openxmldev/linq-to-ooxml";

const mainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/**
 * Gets the w:document root element of the main document part.
 *
 * @param packageRootElement The pkg:package root element of the Flat OPC package.
 * @returns The w:document root element of the main document part.
 */
export function getDocument(packageRootElement: XElement): XElement {
  const mainDocumentPart = packageRootElement
    .elements()
    .single((e) => e.attribute(PKG.contentType)?.value === mainContentType);

  return mainDocumentPart.elements(PKG.xmlData).elements(W.document).single();
}
