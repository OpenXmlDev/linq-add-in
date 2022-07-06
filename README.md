# Linq to XML Demo

This very simple Microsoft Word add-in demonstrates the use of the

- [@openxmldev/linq-to-xml](https://www.npmjs.com/package/@openxmldev/linq-to-xml) and
- [@openxmldev/linq-to-ooxml](https://www.npmjs.com/package/@openxmldev/linq-to-ooxml)

libraries. Those libraries enable pure functional transformations of Office Open XML documents.

## Installing

Ensure that you have a current version of Node.js installed. The add-in was tested with Node.js
16.15.1 LTS.

Open a terminal, go to the desired parent folder and clone this repository:

```
git clone https://github.com/OpenXmlDev/linq-add-in.git
```

Next, `cd` into the `linq-add-in` directory and install the dependencies:

```
npm install
```

## Running

To start the development server, launch Microsoft Word, and sideload the add-in, issue the
following command:

```
npm start
```

Use `npm stop` to stop the development server.

## Code Structure

It all starts with the `App` component, which, among other elements, returns a `DefaultButton` with
an `onClick` handler set to `this.click`.

```html
<DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
  Run
</DefaultButton>
```

Within `Word.run()`, the `click` method calls the asynchronous `transformSelection` function, passing
the actual transformation function `removeDirectFormatting` and the Word request context.

```typescript
click = async () => {
  return Word.run(async (context) => {
    const success = await transformSelection(removeDirectFormatting, context);

    // Certain selections (e.g., one or more table cells, one or more table rows)
    // can't be transformed. Replacing the OOXML of the selected range with the
    // transformed OOXML would mess up the selected range.
    // In this simple example, we don't bother showing a dialog.
    if (!success) console.log("The selected range can't be transformed.");
  }).catch(console.error);
};
```

The `transformSelection` function is a generic "driver" that gets and validates a custom `Selection`
object, calls the actual transformation function (e.g., `removeDirectFormatting`), and inserts the
transformed selection back into the document.

```typescript
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
```

The `removeDirectFormatting` function transforms the selection or, in this case, the `w:document`
root element of the main document part.

```typescript
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
```

Finally, the `removeDirectFormattingTransformation` function is the pure functional transformation
used to transform the `document` `XElment` by stripping all direct formatting and leaving only
paragraph and character styles as well as numbering information.

```typescript
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
```
