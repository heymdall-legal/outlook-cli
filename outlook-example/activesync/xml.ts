export function getFirstTagText(xml: string, tagName: string): string {
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)</${tagName}>`, "i");
  const match = pattern.exec(xml);
  return match ? decodeXml(match[1].trim()) : "";
}

export function getAllTagBlocks(xml: string, tagName: string): string[] {
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)</${tagName}>`, "gi");
  return [...xml.matchAll(pattern)].map((match) => match[1]);
}

export function hasSelfClosingTag(xml: string, tagName: string): boolean {
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?\\s*/>`, "i");
  return pattern.test(xml);
}

export function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
