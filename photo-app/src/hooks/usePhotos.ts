import { useState, useEffect, useCallback } from 'react'
import { getAllPhotos, savePhoto, deletePhoto } from '../db/indexeddb'
import type { Photo } from '../types'

export function usePhotos() {
  const [photos, setPhotos] = useState<Photo[]>([])
  const [loading, setLoading] = useState(true)

  const load = useCallback(async () => {
    setLoading(true)
    const all = await getAllPhotos()
    setPhotos(all)
    setLoading(false)
  }, [])

  useEffect(() => {
    load()
  }, [load])

  const addPhotos = useCallback(async (files: FileList) => {
    for (const file of Array.from(files)) {
      const photo: Omit<Photo, 'url'> = {
        id: crypto.randomUUID(),
        name: file.name,
        blob: file,
        createdAt: Date.now(),
        size: file.size,
      }
      await savePhoto(photo)
    }
    await load()
  }, [load])

  const removePhoto = useCallback(async (id: string) => {
    await deletePhoto(id)
    setPhotos((prev) => prev.filter((p) => p.id !== id))
  }, [])

  return { photos, loading, addPhotos, removePhoto }
}
