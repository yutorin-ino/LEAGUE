import { useRef, useState } from 'react'
import { Gallery } from './components/Gallery'
import { PhotoViewer } from './components/PhotoViewer'
import { usePhotos } from './hooks/usePhotos'
import type { Photo } from './types'
import './App.css'

export default function App() {
  const { photos, loading, addPhotos, removePhoto } = usePhotos()
  const [selected, setSelected] = useState<Photo | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const selectedIndex = selected ? photos.findIndex((p) => p.id === selected.id) : -1

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      await addPhotos(e.target.files)
      e.target.value = ''
    }
  }

  return (
    <div className="app">
      <header className="app-header">
        <h1>フォトアルバム</h1>
        <span className="count">{photos.length} 枚</span>
      </header>

      <main className="app-main">
        {loading ? (
          <div className="loading">読み込み中...</div>
        ) : (
          <Gallery photos={photos} onSelect={setSelected} />
        )}
      </main>

      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        multiple
        hidden
        onChange={handleFileChange}
      />

      <button className="fab" onClick={() => fileInputRef.current?.click()} aria-label="写真を追加">
        ＋
      </button>

      {selected && (
        <PhotoViewer
          photo={selected}
          onClose={() => setSelected(null)}
          onDelete={removePhoto}
          onPrev={selectedIndex > 0 ? () => setSelected(photos[selectedIndex - 1]) : null}
          onNext={selectedIndex < photos.length - 1 ? () => setSelected(photos[selectedIndex + 1]) : null}
        />
      )}
    </div>
  )
}
