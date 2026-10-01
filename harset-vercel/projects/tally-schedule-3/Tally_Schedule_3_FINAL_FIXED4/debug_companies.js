const http = require('http');

const xml = `<ENVELOPE>
    <HEADER>
        <VERSION>1</VERSION>
        <TALLYREQUEST>Export</TALLYREQUEST>
        <TYPE>Collection</TYPE>
        <ID>MyCompanyList</ID>
    </HEADER>
    <BODY>
        <DESC>
            <STATICVARIABLES>
                <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
            <TDL>
                <TDLMESSAGE>
                    <COLLECTION NAME="MyCompanyList">
                        <TYPE>Company</TYPE>
                        <FETCH>Name</FETCH>
                    </COLLECTION>
                </TDLMESSAGE>
            </TDL>
        </DESC>
    </BODY>
</ENVELOPE>`;

const options = {
    hostname: '172.16.1.13',
    port: 9000,
    path: '/',
    method: 'POST',
    headers: {
        'Content-Type': 'text/xml',
        'Content-Length': Buffer.byteLength(xml)
    }
};

const req = http.request(options, (res) => {
    let data = '';
    res.on('data', c => data += c);
    res.on('end', () => {
        console.log('=== RAW RESPONSE ===');
        console.log(data);
        console.log('===================');
        // Try different tag patterns
        const names = data.match(/<NAME>(.*?)<\/NAME>/gi) || [];
        console.log('Matched <NAME> tags:', names);
        const companies = data.match(/<COMPANY>(.*?)<\/COMPANY>/gi) || [];
        console.log('Matched <COMPANY> tags:', companies);
    });
});

req.on('error', e => console.error('ERROR:', e.message));
req.write(xml);
req.end();
