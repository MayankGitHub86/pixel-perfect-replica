export function sanitizeAndExtractText(htmlOrText: string): string {
  if (!htmlOrText) return '';
  if (!htmlOrText.includes('<') && !htmlOrText.includes('>')) {
    return htmlOrText.trim();
  }

  let text = htmlOrText;
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<head[\s\S]*?<\/head>/gi, '');
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<br\s*[\/]?>/gi, '\n');
  text = text.replace(/<p[^>]*>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<div[^>]*>/gi, '\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<li[^>]*>/gi, '\n• ');
  text = text.replace(/<[^>]+>/g, '');
  text = text.replace(/&nbsp;/gi, ' ')
             .replace(/&amp;/gi, '&')
             .replace(/&lt;/gi, '<')
             .replace(/&gt;/gi, '>')
             .replace(/&quot;/gi, '"')
             .replace(/&#39;/gi, "'");
  text = text.split('\n').map(l => l.trim()).filter(Boolean).join('\n');
  return text;
}

const sampleHtml = `<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body><div dir="ltr">name : xyz<div>yshnsghxgshsgxjmasgjytgeuwsdj</div><div>9874563210</div></div></body></html>`;
console.log('Sanitized output:\n' + sanitizeAndExtractText(sampleHtml));
