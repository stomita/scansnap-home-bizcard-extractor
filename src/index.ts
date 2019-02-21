import fs from 'fs';
import os from 'os';
import path from 'path';
import sqlite3 from 'sqlite3';
import jsforce, { Connection } from 'jsforce';

const SCANSNAP_DB_DIR = path.join(os.homedir(), '/Library/Application Support/PFU/ScanSnap Home/Managed/');
const SCANSNAP_FILE_DIR = path.join(os.homedir(), '/Documents/ScanSnapHome.localized');
const SCANSNAP_DB_FILE = path.join(SCANSNAP_DB_DIR, 'ScanSnapHome.sqlite');
const TIMESTAMP_FILE = '.timestamp';

type ContentEntry = {
  ZFAMILYNAME: string | null,
  ZFIRSTNAME: string | null,
  ZCOMPANY: string | null,
  ZDEPARTMENT: string | null,
  ZJOBTITLE: string | null,
  ZZIPCODE: string | null,
  ZADDRESS: string | null,
  ZPHONENUMBER: string | null,
  ZFAXNUMBER: string | null,
  ZEMAIL: string | null,
  ZFILENAME: string,
  ZCREATEDATE: number,
};

type Lead = {
  Id?: string | null,
  FirstName: string | null,
  LastName: string | null,
  Company: string | null,
  Title: string | null,
  PostalCode: string | null,
  State: string | null, 
  Phone: string | null,
  Fax: string | null,
  Email: string | null,
};

async function fetchRecs(db: sqlite3.Database, sql: string) {
  console.log('fetching', sql);
  return new Promise<ContentEntry[]>((resolve, reject) => {
    const recs: ContentEntry[] = [];
    db.each(sql, (err: Error, rec: any) => {
      if (err) { reject(err); }
      recs.push(rec);
    }, (err) => {
      if (err) { reject(err); }
      resolve(recs);
    });
  });
}

async function main() {
  const db = new sqlite3.Database(SCANSNAP_DB_FILE);
  const conn: Connection = (jsforce as any).registry.getConnection(process.env.SF_USERNAME);
  let lastDate: string
  try {
    lastDate = fs.readFileSync(TIMESTAMP_FILE, 'utf8');
  } catch (e) {
    lastDate = '0';
  }
  try {
    const sql = `
    SELECT
      ZFAMILYNAME, ZFIRSTNAME, ZCOMPANY, ZDEPARTMENT, ZJOBTITLE,
      ZZIPCODE, ZADDRESS, ZPHONENUMBER, ZFAXNUMBER, ZEMAIL, ZMEMO,
      ZFILENAME, ZCREATEDATE
    FROM ZCONTENT
    WHERE ZCREATEDATE > ${lastDate}
    ORDER BY ZCREATEDATE DESC
    `;
    const recs = await fetchRecs(db, sql);
    const leads: Lead[] = recs.map((rec) => ({
      FirstName: rec.ZFIRSTNAME,
      LastName: rec.ZFAMILYNAME,
      Company: rec.ZCOMPANY,
      Title: [
        ...(rec.ZDEPARTMENT ? [rec.ZDEPARTMENT] : []),
        ...(rec.ZJOBTITLE ? [rec.ZJOBTITLE] : []),
      ].join(' '),
      PostalCode: rec.ZZIPCODE,
      State: rec.ZADDRESS,
      Phone: rec.ZPHONENUMBER,
      Fax: rec.ZFAXNUMBER,
      Email: rec.ZEMAIL,
    }))
    console.log('creating leads =>', leads);
    const rets = await conn.sobject('Lead').create(leads);
    console.log('results =>', rets);
    if (!Array.isArray(rets)) { return; }
    const retEntries: IterableIterator<[number, jsforce.RecordResult]> = rets.entries();
    for (const [i, ret] of retEntries) {
      if (!ret.success) { continue; }
      const id = ret.id
      const rec = recs[i];
      const email = rec.ZEMAIL;
      const fileName = rec.ZFILENAME;
      const filePath = path.join(SCANSNAP_FILE_DIR, fileName);
      const data = fs.readFileSync(filePath);
      console.log('#####', email, '#####');
      const contentVersion = {
        Title: fileName,
        VersionData: data.toString('base64'),
        PathOnClient: fileName,
      };
      const contentDocumentLink = {
        LinkedEntityId: id,
        ContentDocumentId: '' as string | undefined,
        ShareType: 'V',
      };
      const ret2 = await conn.sobject('ContentVersion').create(contentVersion);
      if (ret2.success) {
        console.log('content version created', ret2.id);
        const version: any = await conn.sobject('ContentVersion').findOne({ Id: ret2.id }, ['Id', 'ContentDocumentId']);
        contentDocumentLink.ContentDocumentId = version.ContentDocumentId;
        await conn.sobject('ContentDocumentLink').create(contentDocumentLink);
      }
    }
    if (recs[0]) {
      fs.writeFileSync(TIMESTAMP_FILE, String(recs[0].ZCREATEDATE), 'utf8');
    }
  } catch(e) {
    console.log(e);
    db.close();
  }
}

main();