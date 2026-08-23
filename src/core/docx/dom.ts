/**
 * The application works on `@xmldom/xmldom`'s DOM (a Node-side XML DOM), not the
 * browser DOM. Re-exporting the types from one place keeps every module honest about
 * which DOM it is using.
 */
export type { Attr, Document, Element, Node, Text } from '@xmldom/xmldom';
