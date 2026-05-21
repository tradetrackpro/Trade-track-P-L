const APP_DATA_FOLDER = 'appDataFolder';
const BACKUP_FILE_NAME = 'tradetrack_backup.json';

export interface BackupData {
  history: any[];
  userName: string;
  timestamp: number;
  totalFunds?: number;
  openingBalance?: number;
  fundTransactions?: any[];
}

export const findBackupFile = async (accessToken: string): Promise<string | null> => {
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?spaces=${APP_DATA_FOLDER}&q=name='${BACKUP_FILE_NAME}' and trashed=false`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error('Failed to search for backup file');
  }

  const data = await response.json();
  return data.files && data.files.length > 0 ? data.files[0].id : null;
};

export const createOrUpdateBackup = async (accessToken: string, backupData: BackupData): Promise<void> => {
  const fileId = await findBackupFile(accessToken);
  const metadata = {
    name: BACKUP_FILE_NAME,
    ...(fileId ? {} : { parents: [APP_DATA_FOLDER] }),
  };

  const fileContent = JSON.stringify(backupData);
  const boundary = '314159265358979323846';
  
  // Construct a standard multipart/related body for Google Drive API V3 upload
  const delimiter = `\r\n--${boundary}\r\n`;
  const close_delim = `\r\n--${boundary}--`;
  
  const body = 
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: application/json\r\n\r\n' +
    fileContent +
    close_delim;

  let url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
  let method = 'POST';

  if (fileId) {
    url = `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart`;
    method = 'PATCH';
  }

  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body,
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Backup error details:', errorText);
    throw new Error(`Failed to save backup to Google Drive: ${response.statusText}`);
  }
};

export const restoreBackup = async (accessToken: string): Promise<BackupData | null> => {
  const fileId = await findBackupFile(accessToken);
  if (!fileId) return null;

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error('Failed to download backup from Google Drive');
  }

  return await response.json();
};
