import { openDB, type IDBPDatabase } from 'idb'
import type { Photo } from '../types'

const DB_NAME = 'photo-app'
const DB_VERSION = 1
const STORE = 'photos'

let db: IDBPDatabase | null = null

async function getDB() {
  if (!db) {
    db = await openDB(DB_NAME, DB_VERSION, {
      upgrade(database) {
        if (!database.objectStoreNames.contains(STORE)) {
          database.createObjectStore(STORE, { keyPath: 'id' })
        }
      },
    })
  }
  return db
}

export async function savePhoto(photo: Omit<Photo, 'url'>): Promise<void> {
  const database = await getDB()
  await database.put(STORE, photo)
}

export async function getAllPhotos(): Promise<Photo[]> {
  const database = await getDB()
  const records = await database.getAll(STORE)
  return records
    .sort((a, b) => b.createdAt - a.createdAt)
    .map((r) => ({ ...r, url: URL.createObjectURL(r.blob) }))
}

export async function deletePhoto(id: string): Promise<void> {
  const database = await getDB()
  await database.delete(STORE, id)
}
