(function () {
  const DELETED_NOTES_KEY = "experimental-note-deleted-notes:v1";
  const HARD_DELETED_NOTE_IDS = new Set(["2026-06-09-thiophene-lithiation-benzaldehyde"]);
  const DELETE_CONFIRM_TEXT = "このノートを削除/アーカイブします。終了登録とは別操作です。よろしいですか？";

  function loadDeletedNoteIds() {
    try {
      return new Set([...HARD_DELETED_NOTE_IDS, ...JSON.parse(localStorage.getItem(DELETED_NOTES_KEY) || "[]")]);
    } catch {
      return new Set(HARD_DELETED_NOTE_IDS);
    }
  }

  function saveDeletedNoteIds(ids) {
    localStorage.setItem(DELETED_NOTES_KEY, JSON.stringify([...ids]));
  }

  function markNoteDeleted(noteId) {
    const deleted = loadDeletedNoteIds();
    deleted.add(noteId);
    saveDeletedNoteIds(deleted);
  }

  function confirmDeleteArchive() {
    return confirm(DELETE_CONFIRM_TEXT);
  }

  window.SharedNoteState = {
    DELETE_CONFIRM_TEXT,
    DELETED_NOTES_KEY,
    loadDeletedNoteIds,
    saveDeletedNoteIds,
    markNoteDeleted,
    confirmDeleteArchive
  };
})();
