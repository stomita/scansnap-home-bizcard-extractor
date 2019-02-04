import fs from 'fs';
import os from 'os';
import path from 'path';
import sqlite3 from 'sqlite3';
import jsforce, { Connection } from 'jsforce';

const SCANSNAP_DB_DIR = path.join(os.homedir(), '/Library/Application Support/PFU/ScanSnap Home/Managed/');
const SCANSNAP_FILE_DIR = path.join(os.homedir(), '/Documents/ScanSnapHomeフォルダ');
const SCANSNAP_DB_FILE = path.join(SCANSNAP_DB_DIR, 'ScanSnapHome.sqlite');

type ContentEntry = {
  ZFILENAME: string,
  ZFAMILYNAME: string,
  ZFIRSTNAME: string,
  ZEMAIL: string,
  ZCOMPANY: string,
  ZADDRESS: string
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
  try {
    const sql = 'SELECT * FROM ZCONTENT';
    const recs = await fetchRecs(db, sql);
    console.log('recs', recs);
    for (const rec of recs) {
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
        LinkedEntityId: '' as string | undefined,
        ContentDocumentId: '' as string | undefined,
        ShareType: 'V',
      };
      const contact = await conn.sobject('Contact').findOne({ Email: email }).sort('CreatedDate', 'DESC');
      if (contact) {
        console.log('Found contact: ', (contact as any).Name);
        contentDocumentLink.LinkedEntityId = contact.Id;
      } else {
        const lead = await conn.sobject('Lead').findOne({ Email: email }).sort('CreatedDate', 'DESC');
        if (lead) {
          console.log('Found lead: ', (lead as any).Name);
          contentDocumentLink.LinkedEntityId = lead.Id;
        }
      }
      if (contentDocumentLink.LinkedEntityId) {
        const ret = await conn.sobject('ContentVersion').create(contentVersion);
        if (ret.success) {
          console.log('content version created', ret.id);
          const version: any = await conn.sobject('ContentVersion').findOne({ Id: ret.id }, ['Id', 'ContentDocumentId']);
          contentDocumentLink.ContentDocumentId = version.ContentDocumentId;
          await conn.sobject('ContentDocumentLink').create(contentDocumentLink);
        }
      }
    }
  } catch(e) {
    console.log(e);
    db.close();
  }
}

main();