import { useEffect } from 'react'
import type { Photo } from '../types'
import styles from './PhotoViewer.module.css'

interface Props {
  photo: Photo
  onClose: () => void
  onDelete: (id: string) => void
  onPrev: (() => void) | null
  onNext: (() => void) | null
}

export function PhotoViewer({ photo, onClose, onDelete, onPrev, onNext }: Props) {
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
      if (e.key === 'ArrowLeft') onPrev?.()
      if (e.key === 'ArrowRight') onNext?.()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose, onPrev, onNext])

  const handleDelete = () => {
    if (confirm(`「${photo.name}」を削除しますか？`)) {
      onDelete(photo.id)
      onClose()
    }
  }

  const formatSize = (bytes: number) => {
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  }

  return (
    <div className={styles.overlay} onClick={onClose}>
      <div className={styles.header} onClick={(e) => e.stopPropagation()}>
        <button className={styles.closeBtn} onClick={onClose}>✕</button>
        <span className={styles.title}>{photo.name}</span>
        <button className={styles.deleteBtn} onClick={handleDelete}>🗑️</button>
      </div>

      <div className={styles.imageWrap} onClick={(e) => e.stopPropagation()}>
        {onPrev && (
          <button className={`${styles.navBtn} ${styles.prev}`} onClick={onPrev}>‹</button>
        )}
        <img src={photo.url} alt={photo.name} className={styles.image} />
        {onNext && (
          <button className={`${styles.navBtn} ${styles.next}`} onClick={onNext}>›</button>
        )}
      </div>

      <div className={styles.footer} onClick={(e) => e.stopPropagation()}>
        <span>{new Date(photo.createdAt).toLocaleDateString('ja-JP')}</span>
        <span>{formatSize(photo.size)}</span>
      </div>
    </div>
  )
}
