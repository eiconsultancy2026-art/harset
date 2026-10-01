const http = require('http');
const fs = require('fs');

const xml = `<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Export Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <EXPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>Group Summary</REPORTNAME>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    <GROUPNAME>Current Assets</GROUPNAME>
                    <SVCURRENTCOMPANY>ECS1</SVCURRENTCOMPANY>
                    <EXPLODEFLAG>Yes</EXPLODEFLAG>
                    <ISDETAILED>Yes</ISDETAILED>
                    <DSPSHOWALLLEDGERS>Yes</DSPSHOWALLLEDGERS>
                    <EXPLODEALLLEVELS>Yes</EXPLODEALLLEVELS>
                </STATICVARIABLES>
            </REQUESTDESC>
        </EXPORTDATA>
    </BODY>
</ENVELOPE>`;

const req = http.request({
    hostname: '172.16.1.13',
    port: 9000,
    path: '/',
    method: 'POST',
    headers: { 'Content-Length': Buffer.byteLength(xml) }
}, (res) => {
    let data = '';
    res.on('data', c => data += c);
    res.on('end', () => { fs.writeFileSync('test_current_assets_ecs1.xml', data); console.log('done'); });
});
req.write(xml);
req.end();
