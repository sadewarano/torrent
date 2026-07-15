// PATCH DI DALAM webtorrent-custom.js
// Cari fungsi Torrent.prototype._select

Torrent.prototype._select = function () {
  const self = this
  if (self.noMorePieces) return

  const segment = self.client.segment || 0; // [BARU] Ambil segment dari client
  const mode = self.client.segmentMode || 'PREFER'; // [BARU] STRICT / PREFER
  const totalPieces = self.pieces.length;
  const segmentCount = self.client.segmentCount || 5; // [BARU] Default 5 browser

  const start = Math.floor(totalPieces / segmentCount * segment);
  const end = Math.floor(totalPieces / segmentCount * (segment + 1));

  // [BARU] Cek apakah segment saya sudah 100%
  let mySegmentDone = true;
  for(let i=start; i<end; i++){
    if(!self.bitfield.get(i)){
      mySegmentDone = false;
      break;
    }
  }
  self._mySegmentDone = mySegmentDone;

  // [BARU] PIECE PICKER DENGAN BOBOT
  let bestPiece = null;
  let bestScore = -1;

  for (let i = 0; i < totalPieces; i++) {
    if (self.bitfield.get(i)) continue; // sudah punya
    if (!self._hasPeerForPiece(i)) continue; // tidak ada peer yg punya

    let rarity = self._rarityMap[i] || 0; // semakin kecil semakin langka
    let score = 1000000 - rarity; // base score: rarest first

    const inMySegment = i >= start && i < end;

    if (inMySegment) {
      score += 100000; // [BOBOT BESAR] Prioritas segment sendiri
    } else if (mode === 'STRICT' &&!mySegmentDone) {
      score = -1; // [STRICT] Abaikan piece luar segment jika belum selesai
    } else if (self._isPieceIdle(i, 60000)) {
      score += 50000000; // [FALLBACK] Piece idle 60s boleh diambil siapa saja
    }

    if (score > bestScore) {
      bestScore = score;
      bestPiece = i;
    }
  }

  if (bestPiece!== null) {
    self._download(bestPiece);
  }
}

// Helper baru
Torrent.prototype._isPieceIdle = function(pieceIndex, timeoutMs){
  // Simpan timestamp terakhir kali piece ini di request
  this._pieceRequestTime = this._pieceRequestTime || {};
  let last = this._pieceRequestTime[pieceIndex] || 0;
  return Date.now() - last > timeoutMs;
}