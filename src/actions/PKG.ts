import { XName, XNamespace } from "@openxmldev/linq-to-xml";

export class PKG {
  public static readonly pkg: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/2006/xmlPackage");
  public static readonly contentType: XName = PKG.pkg.getName("contentType");
  public static readonly xmlData: XName = PKG.pkg.getName("xmlData");
}
