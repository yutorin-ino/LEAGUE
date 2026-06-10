import type { Photo } from '../types'
import styles from './Gallery.module.css'

interface Props {
  photos: Photo[]
  onSelect: (photo: Photo) => void
}

export function Gallery({ photos, onSelect }: Props) {
  if (photos.length === 0) {
    return (
      <div className={styles.empty}>
        <div className={styles.emptyIcon}>🖼️</div>
        <p>写真がありません</p>
        <p className={styles.emptyHint}>右下の＋ボタンから追加してください</p>
      </div>
    )
  }

  return (
    <div className={styles.grid}>
      {photos.map((photo) => (
        <button
          key={photo.id}
          className={styles.cell}
          onClick={() => onSelect(photo)}
          aria-label={photo.name}
        >
          <img src={photo.url} alt={photo.name} loading="lazy" />
        </button>
      ))}
    </div>
  )
}
