import fs from 'fs';

try {
    const content = fs.readFileSync('output.txt', 'utf16le');
    const lines = content.split('\n');
    const relevant = lines.filter(l =>
        l.includes('19.53') ||
        l.includes('chars') ||
        l.includes('Key:') ||
        l.includes('ISSUE')
    );
    fs.writeFileSync('log_summary.txt', relevant.slice(0, 50).join('\n'));
} catch (e) {
    console.error(e);
}
