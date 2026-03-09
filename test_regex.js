const content = '    # 3=Deleted, 4=Outbox, 5=Sent, 16=Drafts, 23=Junk';
const regex = /((?:"(?:\\.|[^"])*?"|'(?:\\.|[^'])*?'))|(<#[\s\S]*?#>|#[^\n]*)/g;
const result = content.replace(regex, (match, p1, p2) => {
  console.log('Match:', JSON.stringify(match));
  console.log('p1:', JSON.stringify(p1));
  console.log('p2:', JSON.stringify(p2));
  if (p1) return p1;
  return '';
});
console.log('Result:', JSON.stringify(result));
