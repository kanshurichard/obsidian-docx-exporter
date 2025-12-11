import {
  App,
  Plugin,
  Notice,
  MarkdownRenderer,
  Modal,
  TFile,
  requestUrl,
  getLanguage,
  Component
} from 'obsidian';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  ExternalHyperlink,
  IBordersOptions,
  BorderStyle,
  ImageRun,
  IParagraphOptions,
  ITableCellOptions
} from 'docx';
import * as JSZip from 'jszip';

// --- 多语言支持 ---
const locales = {
  en: {
    "PLUGIN_NAME": "DOCX Exporter",
    "EXPORT_COMMAND_NAME": "Export current note to DOCX",
    "LOADING_PLUGIN": "Loading DOCX Exporter Plugin",
    "UNLOADING_PLUGIN": "Unloading DOCX Exporter Plugin",
    "NO_ACTIVE_FILE": "No active file to export.",
    "FILE_ALREADY_EXISTS_TITLE": "File already exists",
    "OVERWRITE_CONFIRMATION": "Do you want to overwrite the existing .docx file?",
    "BUTTON_OVERWRITE": "Overwrite",
    "BUTTON_CANCEL": "Cancel",
    "EXPORT_SUCCESSFUL": "Successfully exported to \"{0}\"",
    "EXPORT_FAILED": "Failed to export DOCX. Check developer console for details.",
    "SAVE_FAILED": "Failed to save file: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Failed to download image: {0}",
    "IMAGE_LINK_MISSING": "Image link attribute is missing, skipping export.",
    "LOCAL_IMAGE_NOT_FOUND": "Local image not found: {0}, skipping export.",
    "LOCAL_IMAGE_FILE_EMPTY": "Local image file is empty, skipping export.",
    "UNSUPPORTED_IMAGE_FORMAT": "Unsupported image format for: {0}",
    "IMAGE_DATA_INVALID": "Image data invalid, skipping export.",
    "IMAGE_PROCESSING_ERROR": "Error processing image: {0}",
    "FILE_READ_FAILED": "Failed to read local image file: {0}, check file permissions.",
    "SAVE_FILE_DIALOG_TITLE": "Save DOCX file",
    "SAVE_FILE_INPUT_LABEL": "File path (relative to vault root):",
    "BUTTON_SAVE": "Save",
    "INVALID_FILE_PATH": "Invalid file path. Please enter a valid path.",
    "EXPORTING_START": "Starting DOCX export...",
    "DOWNLOADING_IMAGE": "Downloading image {0} of {1}..."
  },
  zh: {
    "PLUGIN_NAME": "DOCX 导出器",
    "EXPORT_COMMAND_NAME": "导出当前笔记为 DOCX",
    "LOADING_PLUGIN": "正在加载 DOCX 导出插件",
    "UNLOADING_PLUGIN": "正在卸载 DOCX 导出插件",
    "NO_ACTIVE_FILE": "没有活动的笔记可导出。",
    "FILE_ALREADY_EXISTS_TITLE": "文件已存在",
    "OVERWRITE_CONFIRMATION": "是否要覆盖现有的 .docx 文件？",
    "BUTTON_OVERWRITE": "覆盖",
    "BUTTON_CANCEL": "取消",
    "EXPORT_SUCCESSFUL": "成功导出到 “{0}”",
    "EXPORT_FAILED": "导出 DOCX 失败。请检查开发者控制台以获取详细信息。",
    "SAVE_FAILED": "保存文件失败：{0}",
    "DOWNLOAD_IMAGE_FAILED": "下载图片失败：{0}",
    "IMAGE_LINK_MISSING": "图片链接属性缺失，跳过导出。",
    "LOCAL_IMAGE_NOT_FOUND": "未找到本地图片：{0}，跳过导出。",
    "LOCAL_IMAGE_FILE_EMPTY": "本地图片文件为空，跳过导出。",
    "UNSUPPORTED_IMAGE_FORMAT": "不支持的图片格式：{0}",
    "IMAGE_DATA_INVALID": "图片数据无效，跳过导出。",
    "IMAGE_PROCESSING_ERROR": "处理图片时出错：{0}",
    "FILE_READ_FAILED": "读取本地图片文件失败：{0}，请检查文件权限。",
    "SAVE_FILE_DIALOG_TITLE": "保存 DOCX 文件",
    "SAVE_FILE_INPUT_LABEL": "文件路径（相对于库根目录）：",
    "BUTTON_SAVE": "保存",
    "INVALID_FILE_PATH": "无效的文件路径。请输入一个有效的路径。",
    "FILE_SAVE_LOCATION_NOTICE": "已将 DOCX 文件导出到和笔记文件相同的文件夹内。",
    "EXPORTING_START": "开始导出 DOCX...",
    "DOWNLOADING_IMAGE": "正在下载第 {0} 张图片，共 {1} 张..."
  },
  'zh-tw': {
    "PLUGIN_NAME": "DOCX 匯出器",
    "EXPORT_COMMAND_NAME": "匯出目前筆記為 DOCX",
    "LOADING_PLUGIN": "正在載入 DOCX 匯出插件",
    "UNLOADING_PLUGIN": "正在卸載 DOCX 匯出插件",
    "NO_ACTIVE_FILE": "沒有活動的筆記可匯出。",
    "FILE_ALREADY_EXISTS_TITLE": "文件已存在",
    "OVERWRITE_CONFIRMATION": "是否要覆蓋現有的 .docx 文件？",
    "BUTTON_OVERWRITE": "覆蓋",
    "BUTTON_CANCEL": "取消",
    "EXPORT_SUCCESSFUL": "成功匯出到 「{0}」",
    "EXPORT_FAILED": "匯出 DOCX 失敗。請檢查開發者控制台以取得詳細資訊。",
    "SAVE_FAILED": "儲存文件失敗：{0}",
    "DOWNLOAD_IMAGE_FAILED": "下載圖片失敗：{0}",
    "IMAGE_LINK_MISSING": "圖片連結屬性缺失，跳過匯出。",
    "LOCAL_IMAGE_NOT_FOUND": "找不到本地圖片：{0}，跳過匯出。",
    "LOCAL_IMAGE_FILE_EMPTY": "本地圖片文件為空，跳過匯出。",
    "UNSUPPORTED_IMAGE_FORMAT": "不支援的圖片格式：{0}",
    "IMAGE_DATA_INVALID": "圖片數據無效，跳過匯出。",
    "IMAGE_PROCESSING_ERROR": "處理圖片時出錯：{0}",
    "FILE_READ_FAILED": "讀取本地圖片文件失敗：{0}，請檢查文件權限。",
    "SAVE_FILE_DIALOG_TITLE": "儲存 DOCX 文件",
    "SAVE_FILE_INPUT_LABEL": "文件路徑（相對於庫根目錄）：",
    "BUTTON_SAVE": "儲存",
    "INVALID_FILE_PATH": "無效的文件路徑。請輸入一個有效的路徑。",
    "FILE_SAVE_LOCATION_NOTICE": "已將 DOCX 文件匯出到和筆記文件相同的資料夾內。",
    "EXPORTING_START": "開始匯出 DOCX...",
    "DOWNLOADING_IMAGE": "正在下載第 {0} 張圖片，共 {1} 張..."
  },
  ja: {
    "PLUGIN_NAME": "DOCXエクスポート",
    "EXPORT_COMMAND_NAME": "現在のノートをDOCXとしてエクスポート",
    "LOADING_PLUGIN": "DOCXエクスポートプラグインを読み込み中",
    "UNLOADING_PLUGIN": "DOCXエクスポートプラグインをアンロード中",
    "NO_ACTIVE_FILE": "エクスポートするアクティブなファイルがありません。",
    "FILE_ALREADY_EXISTS_TITLE": "ファイルは既に存在します",
    "OVERWRITE_CONFIRMATION": "既存の.docxファイルを上書きしますか？",
    "BUTTON_OVERWRITE": "上書き",
    "BUTTON_CANCEL": "キャンセル",
    "EXPORT_SUCCESSFUL": "「{0}」にエクスポートしました",
    "EXPORT_FAILED": "DOCXのエクスポートに失敗しました。詳細は開発者コンソールを確認してください。",
    "SAVE_FAILED": "ファイルの保存に失敗しました: {0}",
    "DOWNLOAD_IMAGE_FAILED": "画像のダウンロードに失敗しました: {0}",
    "IMAGE_LINK_MISSING": "画像リンク属性がありません。エクスポートをスキップします。",
    "LOCAL_IMAGE_NOT_FOUND": "ローカル画像が見つかりません: {0}。エクスポートをスキップします。",
    "LOCAL_IMAGE_FILE_EMPTY": "ローカル画像ファイルが空です。エクスポートをスキップします。",
    "UNSUPPORTED_IMAGE_FORMAT": "サポートされていない画像形式: {0}",
    "IMAGE_DATA_INVALID": "画像データが無効です。エクスポートをスキップします。",
    "IMAGE_PROCESSING_ERROR": "画像の処理中にエラーが発生しました: {0}",
    "FILE_READ_FAILED": "ローカル画像ファイルの読み取りに失敗しました: {0}。ファイルのアクセス許可を確認してください。",
    "SAVE_FILE_DIALOG_TITLE": "DOCXファイルを保存",
    "SAVE_FILE_INPUT_LABEL": "ファイルパス（ボールトのルートから）:",
    "BUTTON_SAVE": "保存",
    "INVALID_FILE_PATH": "無効なファイルパスです。有効なパスを入力してください。",
    "FILE_SAVE_LOCATION_NOTICE": "DOCXファイルはノートと同じフォルダにエクスポートされました。",
    "EXPORTING_START": "DOCXのエクスポートを開始しています...",
    "DOWNLOADING_IMAGE": "画像 {0}/{1} をダウンロード中..."
  },
  ko: {
    "PLUGIN_NAME": "DOCX 내보내기",
    "EXPORT_COMMAND_NAME": "현재 노트를 DOCX로 내보내기",
    "LOADING_PLUGIN": "DOCX 내보내기 플러그인 로드 중",
    "UNLOADING_PLUGIN": "DOCX 내보내기 플러그인 언로드 중",
    "NO_ACTIVE_FILE": "내보낼 활성 파일이 없습니다.",
    "FILE_ALREADY_EXISTS_TITLE": "파일이 이미 존재합니다",
    "OVERWRITE_CONFIRMATION": "기존 .docx 파일을 덮어쓰시겠습니까?",
    "BUTTON_OVERWRITE": "덮어쓰기",
    "BUTTON_CANCEL": "취소",
    "EXPORT_SUCCESSFUL": "\"{0}\"에 성공적으로 내보냈습니다",
    "EXPORT_FAILED": "DOCX 내보내기 실패. 자세한 내용은 개발자 콘솔을 확인하십시오。",
    "SAVE_FAILED": "파일 저장 실패: {0}",
    "DOWNLOAD_IMAGE_FAILED": "이미지 다운로드 실패: {0}",
    "IMAGE_LINK_MISSING": "이미지 링크 속성이 누락되었습니다. 내보내기를 건너뜁니다.",
    "LOCAL_IMAGE_NOT_FOUND": "로컬 이미지를 찾을 수 없습니다: {0}. 내보내기를 건너뜁니다。",
    "LOCAL_IMAGE_FILE_EMPTY": "로컬 이미지 파일이 비어 있습니다. 내보내기를 건너뜁니다。",
    "UNSUPPORTED_IMAGE_FORMAT": "지원되지 않는 이미지 형식: {0}",
    "IMAGE_DATA_INVALID": "이미지 데이터가 잘못되었습니다. 내보내기를 건너뜁니다。",
    "IMAGE_PROCESSING_ERROR": "이미지 처리 중 오류가 발생했습니다: {0}",
    "FILE_READ_FAILED": "로컬 이미지 파일을 읽는 데 실패했습니다: {0}. 파일 권한을 확인하십시오。",
    "SAVE_FILE_DIALOG_TITLE": "DOCX 파일 저장",
    "SAVE_FILE_INPUT_LABEL": "파일 경로 (볼트 루트 기준):",
    "BUTTON_SAVE": "저장",
    "INVALID_FILE_PATH": "잘못된 파일 경로입니다. 유효한 경로를 입력하십시오。",
    "FILE_SAVE_LOCATION_NOTICE": "DOCX 파일은 노트와 동일한 폴더에 내보내졌습니다。",
    "EXPORTING_START": "DOCX 내보내기를 시작하는 중...",
    "DOWNLOADING_IMAGE": "이미지 다운로드 중 ({0}/{1})..."
  },
  fr: {
    "PLUGIN_NAME": "Exportateur DOCX",
    "EXPORT_COMMAND_NAME": "Exporter la note actuelle en DOCX",
    "LOADING_PLUGIN": "Chargement du plugin d'exportation DOCX",
    "UNLOADING_PLUGIN": "Déchargement du plugin d'exportation DOCX",
    "NO_ACTIVE_FILE": "Aucun fichier actif à exporter.",
    "FILE_ALREADY_EXISTS_TITLE": "Le fichier existe déjà",
    "OVERWRITE_CONFIRMATION": "Voulez-vous écraser le fichier .docx existant ?",
    "BUTTON_OVERWRITE": "Écraser",
    "BUTTON_CANCEL": "Annuler",
    "EXPORT_SUCCESSFUL": "Exporté avec succès vers \"{0}\"",
    "EXPORT_FAILED": "Échec de l'exportation DOCX. Vérifiez la console de développement pour plus de détails.",
    "SAVE_FAILED": "Échec de l'enregistrement du fichier: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Échec du téléchargement de l'image: {0}",
    "IMAGE_LINK_MISSING": "L'attribut de lien d'image est manquant, l'exportation est ignorée.",
    "LOCAL_IMAGE_NOT_FOUND": "Image locale introuvable: {0}, l'exportation est ignorée.",
    "LOCAL_IMAGE_FILE_EMPTY": "Le fichier d'image locale est vide, l'exportation est ignorée.",
    "UNSUPPORTED_IMAGE_FORMAT": "Format d'image non pris en charge pour: {0}",
    "IMAGE_DATA_INVALID": "Données d'image invalides, l'exportation est ignorée.",
    "IMAGE_PROCESSING_ERROR": "Erreur lors du traitement de l'image: {0}",
    "FILE_READ_FAILED": "Échec de la lecture du fichier d'image local: {0}, vérifiez les autorisations du fichier.",
    "SAVE_FILE_DIALOG_TITLE": "Enregistrer le fichier DOCX",
    "SAVE_FILE_INPUT_LABEL": "Chemin du fichier (par rapport à la racine du coffre-fort):",
    "BUTTON_SAVE": "Enregistrer",
    "INVALID_FILE_PATH": "Chemin de fichier invalide. Veuillez entrer un chemin valide.",
    "FILE_SAVE_LOCATION_NOTICE": "Le fichier DOCX a été exporté dans le même dossier que le fichier de note.",
    "EXPORTING_START": "Démarrage de l'exportation DOCX...",
    "DOWNLOADING_IMAGE": "Téléchargement de l'image {0} sur {1}..."
  },
  es: {
    "PLUGIN_NAME": "Exportador DOCX",
    "EXPORT_COMMAND_NAME": "Exportar nota actual a DOCX",
    "LOADING_PLUGIN": "Cargando el plugin de exportación DOCX",
    "UNLOADING_PLUGIN": "Descargando el plugin de exportación DOCX",
    "NO_ACTIVE_FILE": "No hay archivo activo para exportar.",
    "FILE_ALREADY_EXISTS_TITLE": "El archivo ya existe",
    "OVERWRITE_CONFIRMATION": "¿Desea sobrescribir el archivo .docx existente?",
    "BUTTON_OVERWRITE": "Sobrescribir",
    "BUTTON_CANCEL": "Cancelar",
    "EXPORT_SUCCESSFUL": "Exportado con éxito a \"{0}\"",
    "EXPORT_FAILED": "Error al exportar DOCX. Verifique la consola del desarrollador para más detalles.",
    "SAVE_FAILED": "Error al guardar el archivo: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Error al descargar la imagen: {0}",
    "IMAGE_LINK_MISSING": "El atributo de enlace de la imagen falta, se omite la exportación.",
    "LOCAL_IMAGE_NOT_FOUND": "Imagen local no encontrada: {0}, se omite la exportación.",
    "LOCAL_IMAGE_FILE_EMPTY": "El archivo de imagen local está vacío, se omite la exportación.",
    "UNSUPPORTED_IMAGE_FORMAT": "Formato de imagen no compatible para: {0}",
    "IMAGE_DATA_INVALID": "Datos de imagen no válidos, se omite la exportación.",
    "IMAGE_PROCESSING_ERROR": "Error al procesar la imagen: {0}",
    "FILE_READ_FAILED": "Error al leer el archivo de imagen local: {0}, verifique los permisos del archivo.",
    "SAVE_FILE_DIALOG_TITLE": "Guardar archivo DOCX",
    "SAVE_FILE_INPUT_LABEL": "Ruta del archivo (relativa a la bóveda):",
    "BUTTON_SAVE": "Guardar",
    "INVALID_FILE_PATH": "Ruta de archivo inválida. Por favor, introduzca una ruta válida.",
    "FILE_SAVE_LOCATION_NOTICE": "El archivo DOCX ha sido exportado a la misma carpeta que el archivo de notas.",
    "EXPORTING_START": "Iniciando exportación a DOCX...",
    "DOWNLOADING_IMAGE": "Descargando imagen {0} de {1}..."
  },
  ru: {
    "PLUGIN_NAME": "Экспортер DOCX",
    "EXPORT_COMMAND_NAME": "Экспортировать текущую заметку в DOCX",
    "LOADING_PLUGIN": "Загрузка плагина экспорта DOCX",
    "UNLOADING_PLUGIN": "Выгрузка плагина экспорта DOCX",
    "NO_ACTIVE_FILE": "Нет активного файла для экспорта.",
    "FILE_ALREADY_EXISTS_TITLE": "Файл уже существует",
    "OVERWRITE_CONFIRMATION": "Вы хотите перезаписать существующий файл .docx?",
    "BUTTON_OVERWRITE": "Перезаписать",
    "BUTTON_CANCEL": "Отмена",
    "EXPORT_SUCCESSFUL": "Успешно экспортировано в \"{0}\"",
    "EXPORT_FAILED": "Не удалось экспортировать DOCX. Подробности см. в консоли разработчика.",
    "SAVE_FAILED": "Не удалось сохранить файл: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Не удалось загрузить изображение: {0}",
    "IMAGE_LINK_MISSING": "Отсутствует атрибут ссылки на изображение, экспорт пропущен.",
    "LOCAL_IMAGE_NOT_FOUND": "Локальное изображение не найдено: {0}, экспорт пропущен.",
    "LOCAL_IMAGE_FILE_EMPTY": "Локальный файл изображения пуст, экспорт пропущен.",
    "UNSUPPORTED_IMAGE_FORMAT": "Неподдерживаемый формат изображения для: {0}",
    "IMAGE_DATA_INVALID": "Неверные данные изображения, экспорт пропущен.",
    "IMAGE_PROCESSING_ERROR": "Ошибка обработки изображения: {0}",
    "FILE_READ_FAILED": "Не удалось прочитать локальный файл изображения: {0}, проверьте права доступа к файлу.",
    "SAVE_FILE_DIALOG_TITLE": "Сохранить файл DOCX",
    "SAVE_FILE_INPUT_LABEL": "Путь к файлу (относительно корневой папки хранилища):",
    "BUTTON_SAVE": "Сохранить",
    "INVALID_FILE_PATH": "Неверный путь к файлу. Пожалуйста, введите корректный путь.",
    "FILE_SAVE_LOCATION_NOTICE": "Файл DOCX был экспортирован в ту же папку, что и файл заметки.",
    "EXPORTING_START": "Начало экспорта в DOCX...",
    "DOWNLOADING_IMAGE": "Загрузка изображения {0} из {1}..."
  },
  it: {
    "PLUGIN_NAME": "Esportatore DOCX",
    "EXPORT_COMMAND_NAME": "Esporta nota corrente in DOCX",
    "LOADING_PLUGIN": "Caricamento plugin di esportazione DOCX",
    "UNLOADING_PLUGIN": "Scaricamento plugin di esportazione DOCX",
    "NO_ACTIVE_FILE": "Nessun file attivo da esportare.",
    "FILE_ALREADY_EXISTS_TITLE": "Il file esiste già",
    "OVERWRITE_CONFIRMATION": "Vuoi sovrascrivere il file .docx esistente?",
    "BUTTON_OVERWRITE": "Sovrascrivi",
    "BUTTON_CANCEL": "Annulla",
    "EXPORT_SUCCESSFUL": "Esportato con successo in \"{0}\"",
    "EXPORT_FAILED": "Esportazione DOCX fallita. Controlla la console di sviluppo per i dettagli.",
    "SAVE_FAILED": "Salvataggio file fallito: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Download immagine fallito: {0}",
    "IMAGE_LINK_MISSING": "Attributo link immagine mancante, l'esportazione viene saltata.",
    "LOCAL_IMAGE_NOT_FOUND": "Immagine locale non trovata: {0}, l'esportazione viene saltata.",
    "LOCAL_IMAGE_FILE_EMPTY": "File immagine locale vuoto, l'esportazione viene saltata.",
    "UNSUPPORTED_IMAGE_FORMAT": "Formato immagine non supportato per: {0}",
    "IMAGE_DATA_INVALID": "Dati immagine non validi, l'esportazione viene saltata.",
    "IMAGE_PROCESSING_ERROR": "Errore durante l'elaborazione dell'immagine: {0}",
    "FILE_READ_FAILED": "Lettura file immagine locale fallita: {0}, controlla i permessi del file.",
    "SAVE_FILE_DIALOG_TITLE": "Salva file DOCX",
    "SAVE_FILE_INPUT_LABEL": "Percorso file (relativo alla root del vault):",
    "BUTTON_SAVE": "Salva",
    "INVALID_FILE_PATH": "Percorso file non valido. Inserisci un percorso valido.",
    "FILE_SAVE_LOCATION_NOTICE": "Il file DOCX è stato esportato nella stessa cartella del file di nota.",
    "EXPORTING_START": "Avvio esportazione DOCX...",
    "DOWNLOADING_IMAGE": "Download immagine {0} di {1}..."
  },
  pt: {
    "PLUGIN_NAME": "Exportador DOCX",
    "EXPORT_COMMAND_NAME": "Exportar nota atual para DOCX",
    "LOADING_PLUGIN": "Carregando plugin de exportação DOCX",
    "UNLOADING_PLUGIN": "Descarregando plugin de exportação DOCX",
    "NO_ACTIVE_FILE": "Nenhum arquivo ativo para exportar.",
    "FILE_ALREADY_EXISTS_TITLE": "O arquivo já existe",
    "OVERWRITE_CONFIRMATION": "Deseja sobrescrever o arquivo .docx existente?",
    "BUTTON_OVERWRITE": "Sobrescrever",
    "BUTTON_CANCEL": "Cancelar",
    "EXPORT_SUCCESSFUL": "Exportado com sucesso para \"{0}\"",
    "EXPORT_FAILED": "Falha ao exportar DOCX. Verifique o console do desenvolvedor para detalhes.",
    "SAVE_FAILED": "Falha ao salvar arquivo: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Falha ao baixar imagem: {0}",
    "IMAGE_LINK_MISSING": "Atributo de link de imagem ausente, pulando exportação.",
    "LOCAL_IMAGE_NOT_FOUND": "Imagem local não encontrada: {0}, pulando exportação.",
    "LOCAL_IMAGE_FILE_EMPTY": "Arquivo de imagem local vazio, pulando exportação.",
    "UNSUPPORTED_IMAGE_FORMAT": "Formato de imagem não suportado para: {0}",
    "IMAGE_DATA_INVALID": "Dados de imagem inválidos, pulando exportação.",
    "IMAGE_PROCESSING_ERROR": "Erro ao processar imagem: {0}",
    "FILE_READ_FAILED": "Falha ao ler arquivo de imagem local: {0}, verifique as permissões do arquivo.",
    "SAVE_FILE_DIALOG_TITLE": "Salvar arquivo DOCX",
    "SAVE_FILE_INPUT_LABEL": "Caminho do arquivo (relativo à raiz do vault):",
    "BUTTON_SAVE": "Salvar",
    "INVALID_FILE_PATH": "Caminho de arquivo inválido. Por favor, insira um caminho válido.",
    "FILE_SAVE_LOCATION_NOTICE": "O arquivo DOCX foi exportado para a mesma pasta do arquivo de nota.",
    "EXPORTING_START": "Iniciando exportação DOCX...",
    "DOWNLOADING_IMAGE": "Baixando imagem {0} de {1}..."
  },
  tr: {
    "PLUGIN_NAME": "DOCX Dışa Aktarıcı",
    "EXPORT_COMMAND_NAME": "Mevcut notu DOCX olarak dışa aktar",
    "LOADING_PLUGIN": "DOCX dışa aktarma eklentisi yükleniyor",
    "UNLOADING_PLUGIN": "DOCX dışa aktarma eklentisi kaldırılıyor",
    "NO_ACTIVE_FILE": "Dışa aktarılacak aktif dosya yok.",
    "FILE_ALREADY_EXISTS_TITLE": "Dosya zaten mevcut",
    "OVERWRITE_CONFIRMATION": "Mevcut .docx dosyasının üzerine yazmak istiyor musunuz?",
    "BUTTON_OVERWRITE": "Üzerine yaz",
    "BUTTON_CANCEL": "İptal",
    "EXPORT_SUCCESSFUL": "Başarıyla \"{0}\" konumuna dışa aktarıldı",
    "EXPORT_FAILED": "DOCX dışa aktarma başarısız. Ayrıntılar için geliştirici konsolunu kontrol edin.",
    "SAVE_FAILED": "Dosya kaydetme başarısız: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Resim indirme başarısız: {0}",
    "IMAGE_LINK_MISSING": "Resim bağlantı özniteliği eksik, dışa aktarma atlanıyor.",
    "LOCAL_IMAGE_NOT_FOUND": "Yerel resim bulunamadı: {0}, dışa aktarma atlanıyor.",
    "LOCAL_IMAGE_FILE_EMPTY": "Yerel resim dosyası boş, dışa aktarma atlanıyor.",
    "UNSUPPORTED_IMAGE_FORMAT": "Desteklenmeyen resim formatı: {0}",
    "IMAGE_DATA_INVALID": "Geçersiz resim verisi, dışa aktarma atlanıyor.",
    "IMAGE_PROCESSING_ERROR": "Resim işleme hatası: {0}",
    "FILE_READ_FAILED": "Yerel resim dosyası okuma başarısız: {0}, dosya izinlerini kontrol edin.",
    "SAVE_FILE_DIALOG_TITLE": "DOCX dosyasını kaydet",
    "SAVE_FILE_INPUT_LABEL": "Dosya yolu (kasa kök dizinine göre):",
    "BUTTON_SAVE": "Kaydet",
    "INVALID_FILE_PATH": "Geçersiz dosya yolu. Lütfen geçerli bir yol girin.",
    "FILE_SAVE_LOCATION_NOTICE": "DOCX dosyası, not dosyasıyla aynı klasöre aktarıldı.",
    "EXPORTING_START": "DOCX dışa aktarma başlatılıyor...",
    "DOWNLOADING_IMAGE": "Resim indiriliyor {0}/{1}..."
  },
  de: {
    "PLUGIN_NAME": "DOCX Exportierer",
    "EXPORT_COMMAND_NAME": "Aktuelle Notiz als DOCX exportieren",
    "LOADING_PLUGIN": "DOCX Export-Plugin wird geladen",
    "UNLOADING_PLUGIN": "DOCX Export-Plugin wird entladen",
    "NO_ACTIVE_FILE": "Keine aktive Datei zum Exportieren.",
    "FILE_ALREADY_EXISTS_TITLE": "Datei existiert bereits",
    "OVERWRITE_CONFIRMATION": "Möchten Sie die bestehende .docx-Datei überschreiben?",
    "BUTTON_OVERWRITE": "Überschreiben",
    "BUTTON_CANCEL": "Abbrechen",
    "EXPORT_SUCCESSFUL": "Erfolgreich exportiert nach \"{0}\"",
    "EXPORT_FAILED": "DOCX-Export fehlgeschlagen. Prüfen Sie die Entwicklerkonsole für Details.",
    "SAVE_FAILED": "Fehler beim Speichern der Datei: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Fehler beim Herunterladen des Bildes: {0}",
    "IMAGE_LINK_MISSING": "Bild-Link-Attribut fehlt, Export wird übersprungen.",
    "LOCAL_IMAGE_NOT_FOUND": "Lokales Bild nicht gefunden: {0}, Export wird übersprungen.",
    "LOCAL_IMAGE_FILE_EMPTY": "Lokale Bilddatei ist leer, Export wird übersprungen.",
    "UNSUPPORTED_IMAGE_FORMAT": "Nicht unterstütztes Bildformat für: {0}",
    "IMAGE_DATA_INVALID": "Ungültige Bilddaten, Export wird übersprungen.",
    "IMAGE_PROCESSING_ERROR": "Fehler bei der Bildverarbeitung: {0}",
    "FILE_READ_FAILED": "Fehler beim Lesen der lokalen Bilddatei: {0}, überprüfen Sie die Dateiberechtigungen.",
    "SAVE_FILE_DIALOG_TITLE": "DOCX-Datei speichern",
    "SAVE_FILE_INPUT_LABEL": "Dateipfad (relativ zum Vault-Root):",
    "BUTTON_SAVE": "Speichern",
    "INVALID_FILE_PATH": "Ungültiger Dateipfad. Bitte geben Sie einen gültigen Pfad ein.",
    "FILE_SAVE_LOCATION_NOTICE": "Die DOCX-Datei wurde im gleichen Ordner wie die Notizdatei exportiert.",
    "EXPORTING_START": "DOCX-Export wird gestartet...",
    "DOWNLOADING_IMAGE": "Lade Bild {0} von {1}..."
  },
  ar: {
    "PLUGIN_NAME": "مصدّر DOCX",
    "EXPORT_COMMAND_NAME": "تصدير الملاحظة الحالية إلى DOCX",
    "LOADING_PLUGIN": "جاري تحميل إضافة تصدير DOCX",
    "UNLOADING_PLUGIN": "جاري إلغاء تحميل إضافة تصدير DOCX",
    "NO_ACTIVE_FILE": "لا يوجد ملف نشط للتصدير.",
    "FILE_ALREADY_EXISTS_TITLE": "الملف موجود بالفعل",
    "OVERWRITE_CONFIRMATION": "هل تريد استبدال ملف .docx الموجود؟",
    "BUTTON_OVERWRITE": "استبدال",
    "BUTTON_CANCEL": "إلغاء",
    "EXPORT_SUCCESSFUL": "تم التصدير بنجاح إلى \"{0}\"",
    "EXPORT_FAILED": "فشل تصدير DOCX. راجع وحدة التحكم للمطورين للحصول على التفاصيل.",
    "SAVE_FAILED": "فشل حفظ الملف: {0}",
    "DOWNLOAD_IMAGE_FAILED": "فشل تحميل الصورة: {0}",
    "IMAGE_LINK_MISSING": "سمة رابط الصورة مفقودة، يتم تخطي التصدير.",
    "LOCAL_IMAGE_NOT_FOUND": "الصورة المحلية غير موجودة: {0}، يتم تخطي التصدير.",
    "LOCAL_IMAGE_FILE_EMPTY": "ملف الصورة المحلي فارغ، يتم تخطي التصدير.",
    "UNSUPPORTED_IMAGE_FORMAT": "تنسيق صورة غير مدعوم لـ: {0}",
    "IMAGE_DATA_INVALID": "بيانات الصورة غير صالحة، يتم تخطي التصدير.",
    "IMAGE_PROCESSING_ERROR": "خطأ في معالجة الصورة: {0}",
    "FILE_READ_FAILED": "فشل قراءة ملف الصورة المحلي: {0}، تحقق من أذونات الملف.",
    "SAVE_FILE_DIALOG_TITLE": "حفظ ملف DOCX",
    "SAVE_FILE_INPUT_LABEL": "مسار الملف (نسبة إلى جذر الخزينة):",
    "BUTTON_SAVE": "حفظ",
    "INVALID_FILE_PATH": "مسار ملف غير صالح. الرجاء إدخال مسار صالح.",
    "FILE_SAVE_LOCATION_NOTICE": "تم تصدير ملف DOCX إلى نفس مجلد ملف الملاحظة.",
    "EXPORTING_START": "جاري بدء تصدير DOCX...",
    "DOWNLOADING_IMAGE": "جاري تحميل الصورة {0} من {1}..."
  }
};

class I18N {
  private lang: string;
  private translations: Record<string, string>;

  constructor(app: App) {
    let lang = getLanguage();
    if (!lang) {
      lang = 'en';
    }

    if (locales[lang]) {
      this.lang = lang;
    } else if (lang.includes('-')) {
      const baseLang = lang.split('-')[0];
      this.lang = locales[baseLang] ? baseLang : 'en';
    } else {
      this.lang = 'en';
    }
    this.translations = locales[this.lang] || locales['en'];
  }

  t(key: string, ...args: string[]): string {
    let text = this.translations[key] || locales['en'][key] || key;
    if (args) {
      args.forEach((arg, index) => {
        text = text.replace(new RegExp(`\\{${index}\\}`, 'g'), arg);
      });
    }
    return text;
  }
}

// --- 文件覆盖确认弹窗 ---
class OverwriteConfirmModal extends Modal {
  onConfirm: () => void;
  i18n: I18N;
  constructor(app: App, onConfirm: () => void, i18n: I18N) {
    super(app);
    this.onConfirm = onConfirm;
    this.i18n = i18n;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.createEl("h2", { text: this.i18n.t("FILE_ALREADY_EXISTS_TITLE") });
    contentEl.createEl("p", { text: this.i18n.t("OVERWRITE_CONFIRMATION") });
    const buttonContainer = contentEl.createDiv({ cls: "modal-button-container" });
    buttonContainer.createEl("button", { text: this.i18n.t("BUTTON_OVERWRITE"), cls: "mod-cta" }).addEventListener("click", () => { this.close(); this.onConfirm(); });
    buttonContainer.createEl("button", { text: this.i18n.t("BUTTON_CANCEL") }).addEventListener("click", () => { this.close(); });
  }
  onClose() { let { contentEl } = this; contentEl.empty(); }
}

// --- 插件主类 ---
export default class DocxExporterPlugin extends Plugin {
  i18n: I18N;
  private numberingCounter = 1; // 有序列表编号计数器
  private numberingReferences = new Set<string>(); // 记录所有编号 reference

  private totalNetworkImages = 0;
  private currentImageIndex = 0;

  async onload() {
    this.i18n = new I18N(this.app);
    this.addRibbonIcon('file-output', this.i18n.t("EXPORT_COMMAND_NAME"), () => this.exportCurrentNoteToDocx());
    this.addCommand({
      id: 'export-to-docx',
      name: this.i18n.t("EXPORT_COMMAND_NAME"),
      checkCallback: (checking: boolean) => {
        const activeFile = this.app.workspace.getActiveFile();
        if (activeFile) {
          if (!checking) {
            this.exportCurrentNoteToDocx();
          }
          return true;
        }
        return false;
      }
    });
  }
  onunload() {
  }

  // px转为docx字体单位
  private pxToHalfPoints(px: string): number | undefined {
    const pxValue = parseFloat(px);
    return isNaN(pxValue) ? undefined : Math.round(pxValue * 1.5);
  }

  // rgb颜色转16进制
  private rgbToHex(rgb: string): string | undefined {
    if (!rgb?.startsWith('rgb') || rgb === 'rgba(0, 0, 0, 0)') return undefined;
    const parts = rgb.match(/^rgb(?:a)?\((\d+),\s*(\d+),\s*(\d+)/);
    return parts ? [1, 2, 3].map(i => ('0' + parseInt(parts[i]).toString(16)).slice(-2)).join('') : undefined;
  }

  // base64转ArrayBuffer
  private base64ToArrayBuffer(base64: string): ArrayBuffer {
    const bin = window.atob(base64), len = bin.length, bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }

  // 解析图片尺寸
  private getImageDimensionsFromBuffer(buffer: ArrayBuffer): { width: number, height: number } | null {
    const view = new DataView(buffer);
    try {
      if (view.getUint16(0) === 0xFFD8) { // JPEG
        let offset = 2;
        while (offset < buffer.byteLength) {
          if (view.getUint8(offset) !== 0xFF) { offset++; continue; }
          const marker = view.getUint8(offset + 1);
          if (marker === 0xC0 || marker === 0xC2)
            return { width: view.getUint16(offset + 7), height: view.getUint16(offset + 5) };
          offset += 2 + view.getUint16(offset + 2);
        }
      } else if (view.getUint32(0) === 0x89504E47) { // PNG
        return { width: view.getUint32(16), height: view.getUint32(20) };
      }
    } catch { }
    return null;
  }

  // 检测图片mime
  private detectMimeFromHeader(u8: Uint8Array): string | null {
    if (!u8 || u8.length < 8) return null;
    if (u8[0] === 0x89 && u8[1] === 0x50) return 'image/png';
    if (u8[0] === 0xFF && u8[1] === 0xD8) return 'image/jpeg';
    if (u8[0] === 0x47 && u8[1] === 0x49) return 'image/gif';
    if (u8[0] === 0x42 && u8[1] === 0x4D) return 'image/bmp';
    if (u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 &&
      u8[8] === 0x57 && u8[9] === 0x45) return 'image/webp';
    if (new TextDecoder().decode(u8.slice(0, 64)).includes('<svg')) return 'image/svg+xml';
    return null;
  }

  // mime转扩展名
  private extFromMime(mime: string | null): string | null {
    if (!mime) return null;
    if (mime === 'image/png') return 'png';
    if (mime === 'image/jpeg') return 'jpg';
    if (mime === 'image/gif') return 'gif';
    if (mime === 'image/bmp') return 'bmp';
    if (mime === 'image/svg+xml') return 'svg';
    if (mime === 'image/webp') return 'webp';
    return null;
  }

  // 转义正则
  private escapeRegExp(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  // 解析代码高亮
  private parseSyntaxHighlightedCode(codeElement: HTMLElement, defaultFont: any, defaultSize: number): TextRun[] {
    const runs: TextRun[] = [];
    const codeFont = { name: 'Courier New' };
    Array.from(codeElement.childNodes).forEach(node => {
      if (node.nodeType === Node.TEXT_NODE && node.textContent) {
        node.textContent.split('\n').forEach((seg, idx, arr) => {
          if (seg) runs.push(new TextRun({ text: seg, font: codeFont, color: this.rgbToHex(window.getComputedStyle(codeElement).color), size: defaultSize }));
          if (idx < arr.length - 1) runs.push(new TextRun({ break: 1 }));
        });
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        el.textContent?.split('\n').forEach((seg, idx, arr) => {
          if (seg) runs.push(new TextRun({ text: seg, font: codeFont, color: this.rgbToHex(window.getComputedStyle(el).color), size: defaultSize }));
          if (idx < arr.length - 1) runs.push(new TextRun({ break: 1 }));
        });
      }
    });
    return runs;
  }

  // 解析列表（支持多级缩进和有序/无序）
  private async parseListElement(
    listEl: HTMLUListElement | HTMLOListElement,
    level: number,
    bodyBgColor: string,
    sourcePath: string,
    numberingRef?: string
  ): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    const listType = listEl.tagName === 'OL' ? 'number' : 'bullet';
    let currentNumberingRef = numberingRef;
    if (listType === 'number' && !numberingRef) {
      currentNumberingRef = `default-numbering-${this.numberingCounter++}`;
      this.numberingReferences.add(currentNumberingRef);
    }
    for (const li of Array.from(listEl.children).filter(c => c.tagName === 'LI')) {
      const liElement = li as HTMLLIElement;
      const contentContainer = document.createElement('div');
      let nestedList: HTMLUListElement | HTMLOListElement | null = null;
      for (const child of Array.from(liElement.childNodes)) {
        if (child.nodeType === Node.ELEMENT_NODE && (child.nodeName === 'UL' || child.nodeName === 'OL'))
          nestedList = child as HTMLUListElement | HTMLOListElement;
        else contentContainer.appendChild(child.cloneNode(true));
      }
      if (contentContainer.textContent?.trim() || contentContainer.querySelector('img')) {
        const style = window.getComputedStyle(liElement);
        const paragraphOptions: any = {};
        const bgColor = this.rgbToHex(style.backgroundColor);
        if (bgColor && bgColor !== bodyBgColor)
          paragraphOptions.shading = { type: ShadingType.CLEAR, fill: bgColor, color: "auto" };
        paragraphOptions.indent = { left: 720 * level };
        const children = await this.parseInlineElements(contentContainer, sourcePath);
        if (children.length > 0) {
          const props: any = { ...paragraphOptions, children, spacing: { after: 100 } };
          if (listType === 'bullet') props.bullet = { level };
          else props.numbering = { reference: currentNumberingRef, level };
          paragraphs.push(new Paragraph(props));
        }
      }
      if (nestedList)
        paragraphs.push(...await this.parseListElement(nestedList, level + 1, bodyBgColor, sourcePath, currentNumberingRef));
    }
    return paragraphs;
  }

  // 解析表格
  private async parseTableElement(tableEl: HTMLElement, bodyBgColor: string, sourcePath: string): Promise<Table> {
    const rows: TableRow[] = [];
    let isFirstRow = true;  // 标记是否为表头行

    // 遍历表格行
    for (const row of Array.from(tableEl.querySelectorAll('tr'))) {
      const cells: TableCell[] = [];

      // 遍历单元格
      for (const cell of Array.from(row.querySelectorAll('th, td'))) {
        const cellStyle = window.getComputedStyle(cell);
        const isHeader = cell.tagName.toUpperCase() === 'TH' || isFirstRow;

        // 处理单元格内容
        const inlineElements = await this.parseInlineElements(cell, sourcePath);
        const paragraph = new Paragraph({
          children: inlineElements,
          alignment: this.getCellAlignment(cellStyle.textAlign),
          spacing: { before: 100, after: 100 },
          font: { name: 'Times New Roman' }
        });

        // 为表头单元格添加特殊样式
        cells.push(new TableCell({
          children: [paragraph],
          margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100
          },
          ...(isHeader && {
            shading: {
              fill: this.rgbToHex(cellStyle.backgroundColor) || "E7E6E6",
              type: ShadingType.CLEAR,
              color: "auto"
            },
            verticalAlign: "center"
          }),
          verticalAlign: "center"
        }));
      }

      rows.push(new TableRow({
        children: cells,
        tableHeader: isFirstRow // 标记表头行
      }));
      isFirstRow = false;
    }

    // 创建表格，添加边框样式
    return new Table({
      rows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      margins: { top: 100, bottom: 100 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
        left: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
        right: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "auto" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "auto" }
      }
    });
  }

  // 辅助函数：获取单元格对齐方式
  private getCellAlignment(textAlign: string): AlignmentType {
    switch (textAlign) {
      case 'center': return AlignmentType.CENTER;
      case 'right': return AlignmentType.RIGHT;
      case 'justify': return AlignmentType.JUSTIFIED;
      default: return AlignmentType.LEFT;
    }
  }

  // 解析行内元素（支持超链接、加粗、斜体、代码、图片等）
  private async parseInlineElements(element: HTMLElement, sourcePath: string): Promise<(TextRun | ExternalHyperlink | ImageRun)[]> {
    const runs: (TextRun | ExternalHyperlink | ImageRun)[] = [];
    const mainFont = { name: 'Times New Roman' };
    const codeFont = { name: 'Courier New' };
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        const textContent = child.textContent ?? '';
        if (!textContent.trim()) continue;
        const style = window.getComputedStyle(element);
        runs.push(new TextRun({
          text: textContent,
          color: this.rgbToHex(style.color),
          size: this.pxToHalfPoints(style.fontSize),
          font: mainFont,
          bold: parseInt(style.fontWeight) >= 600,
          italics: style.fontStyle === 'italic'
        }));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        const style = window.getComputedStyle(el);
        // 行内代码特殊处理
        if (el.tagName.toUpperCase() === 'CODE') {
          runs.push(new TextRun({
            text: el.textContent ?? '',
            style: "SourceCode",
            shading: { type: ShadingType.CLEAR, fill: this.rgbToHex(style.backgroundColor) || 'D3D3D3', color: "auto" },
            font: codeFont,
            color: this.rgbToHex(style.color),
            size: this.pxToHalfPoints(style.fontSize)
          }));
          continue;
        }
        const runOptions: any = { text: el.textContent?.trim() || '', color: this.rgbToHex(style.color), size: this.pxToHalfPoints(style.fontSize), font: mainFont };
        switch (el.tagName.toUpperCase()) {
          case 'A':
            runs.push(new ExternalHyperlink({ link: el.getAttribute('href') || '', children: [new TextRun({ text: el.textContent?.trim() || '', style: "Hyperlink", color: this.rgbToHex(style.color) || '0563C1', underline: {} })] }));
            break;
          case 'DEL': runOptions.strike = true; runs.push(new TextRun(runOptions)); break;
          case 'STRONG':
          case 'B': runOptions.bold = true; runs.push(new TextRun(runOptions)); break;
          case 'EM':
          case 'I': runOptions.italics = true; runs.push(new TextRun(runOptions)); break;
          case 'U': runOptions.underline = {}; runs.push(new TextRun(runOptions)); break;
          case 'IMG':
            const imageRun = await this.createImageRun(el as HTMLImageElement, sourcePath);
            if (imageRun) runs.push(imageRun);
            break;
          case 'BR':
            runs.push(new TextRun({ break: 1 }));
            break;
          default:
            if (el.tagName.toUpperCase() === 'SPAN' && el.classList.contains('internal-embed')) {
              const imgEl = el.querySelector('img');
              if (imgEl) {
                const imageRun = await this.createImageRun(imgEl as HTMLImageElement, sourcePath);
                if (imageRun) runs.push(imageRun);
              }
            } else if (runOptions.text && runOptions.text.trim().length > 0) {
              runOptions.text = runOptions.text.trim();
              runs.push(new TextRun(runOptions));
            }
            break;
        }
      }
    }
    return runs;
  }

  private async createImageRun(imgEl: HTMLImageElement, sourcePath: string): Promise<ImageRun | null> {
    const src = imgEl.getAttribute('src');
    const altText = imgEl.getAttribute('alt');
    let buffer: ArrayBuffer | null = null;
    let imageExtension: string | null = null;
    let pathForNotice: string = 'unknown path';

    const isLocalPath = src.startsWith('app://') || src.startsWith('capacitor://');

    try {
      if (src.startsWith('http')) {
        pathForNotice = src;
        this.currentImageIndex++;
        new Notice(this.i18n.t("DOWNLOADING_IMAGE", this.currentImageIndex.toString(), this.totalNetworkImages.toString()));
        const response = await requestUrl({ url: src, method: 'GET' });
        if (response.status !== 200) {
          new Notice(this.i18n.t("DOWNLOAD_IMAGE_FAILED", src));
          return null;
        }
        buffer = response.arrayBuffer;
        imageExtension = src.split('.').pop()?.toLowerCase() || 'jpeg';
      } else if (isLocalPath) {
        const parentSpan = imgEl.parentElement as HTMLElement;
        const pathFromEmbed = parentSpan.getAttribute('alt') || parentSpan.getAttribute('data-href') || parentSpan.getAttribute('data-src');
        if (!pathFromEmbed) {
          new Notice(this.i18n.t("IMAGE_LINK_MISSING"));
          return null;
        }
        pathForNotice = pathFromEmbed;
        const file = this.app.metadataCache.getFirstLinkpathDest(pathFromEmbed, sourcePath);

        if (!file || !(file instanceof TFile)) {
          new Notice(this.i18n.t("LOCAL_IMAGE_NOT_FOUND", pathFromEmbed));
          return null;
        }

        try {
          buffer = await this.app.vault.readBinary(file);
          imageExtension = file.extension;
        } catch (readError) {
          new Notice(this.i18n.t("FILE_READ_FAILED", file.path));
          return null;
        }

      } else if (src.startsWith('data:image')) {
        pathForNotice = 'Base64 embedded image';
        const base64String = src.split(',')[1];
        buffer = this.base64ToArrayBuffer(base64String);
        const mimeType = src.match(/data:image\/(.*?);/)?.[1];
        imageExtension = mimeType || 'png';
      } else {
        new Notice(this.i18n.t("UNSUPPORTED_IMAGE_FORMAT", src));
        return null;
      }

      if (!buffer || !imageExtension) {
        new Notice(this.i18n.t("IMAGE_DATA_INVALID"));
        return null;
      }

      let width = 550;
      let height = 300;
      const dimensionsFromBuffer = this.getImageDimensionsFromBuffer(buffer);
      let naturalWidth = dimensionsFromBuffer?.width;
      let naturalHeight = dimensionsFromBuffer?.height;

      if (!naturalWidth || !naturalHeight) {
        naturalWidth = imgEl.naturalWidth > 0 ? imgEl.naturalWidth : undefined;
        naturalHeight = imgEl.naturalHeight > 0 ? imgEl.naturalHeight : undefined;
      }

      const styleWidth = parseFloat(imgEl.style.width) || imgEl.width;
      const styleHeight = parseFloat(imgEl.style.height) || imgEl.height;

      let finalWidth = styleWidth || naturalWidth;
      let finalHeight = styleHeight || (finalWidth && naturalWidth ? (finalWidth / naturalWidth) * naturalHeight : undefined);

      if (!finalWidth || !finalHeight || finalWidth <= 0 || finalHeight <= 0) {
        finalWidth = width;
        finalHeight = height;
      } else {
        finalWidth = Math.round(finalWidth);
        finalHeight = Math.round(finalHeight);

        if (finalWidth > width) {
          finalHeight = Math.round((width / finalWidth) * finalHeight);
          finalWidth = width;
        }
      }

      return new ImageRun({
        data: buffer,
        transformation: {
          width: finalWidth,
          height: finalHeight
        }
      });

    } catch (error) {
      const displayPath = pathForNotice.length > 50 ? `${pathForNotice.substring(0, 50)}...` : pathForNotice;
      new Notice(this.i18n.t("IMAGE_PROCESSING_ERROR", displayPath));
      return null;
    }
  }

  private countNetworkImages(element: HTMLElement): number {
    let count = 0;
    const images = element.querySelectorAll('img');
    images.forEach(img => {
      if (img.src.startsWith('http')) {
        count++;
      }
    });
    return count;
  }

  private async htmlToDocxObjects(
    element: HTMLElement,
    bodyBgColor: string,
    isTopLevel: boolean,
    indentLevel: number,
    sourcePath: string
  ): Promise<(Paragraph | Table)[]> {
    const docxObjects: (Paragraph | Table)[] = [];
    const children = Array.from(element.childNodes);
    const paragraphStyles: any = {};
    const style = window.getComputedStyle(element);
    switch (style.textAlign) {
      case 'center': paragraphStyles.alignment = AlignmentType.CENTER; break;
      case 'right': paragraphStyles.alignment = AlignmentType.RIGHT; break;
      case 'justify': paragraphStyles.alignment = AlignmentType.JUSTIFIED; break;
    }
    const mainFont = { name: 'Times New Roman' };
    const codeFont = { name: 'Courier New' };

    // 如果没有子元素或只有空文本，则直接返回一个空段落
    if (!children.some(c => c.nodeType === Node.ELEMENT_NODE) && element.textContent?.trim()) {
      const inlineChildren = await this.parseInlineElements(element, sourcePath);
      if (inlineChildren.length > 0) { return [new Paragraph({ children: inlineChildren, font: mainFont })]; }
      return [];
    }

    for (let i = 0; i < children.length; i++) {
      const child = children[i];

      // 修正：确保 child 是一个有效的元素节点
      if (!(child instanceof HTMLElement)) continue;

      const el = child as HTMLElement;
      const tagName = el.tagName?.toUpperCase();
      if (!tagName) continue;

      let currentParagraphOptions: any = { ...paragraphStyles, font: mainFont };

      if (tagName.startsWith('H') && i > 0) {
        const prevChild = children[i - 1] as HTMLElement;
        if (prevChild && prevChild.tagName?.toUpperCase() !== 'HR' && prevChild.tagName?.toUpperCase() !== 'TABLE') {
          docxObjects.push(new Paragraph({}));
        }
      }

      switch (tagName) {
        case 'H1':
        case 'H2':
        case 'H3':
        case 'H4':
        case 'H5':
        case 'H6':
          currentParagraphOptions.heading = HeadingLevel[tagName as keyof typeof HeadingLevel];
          currentParagraphOptions.spacing = { after: 150 };
          const headingChildren = await this.parseInlineElements(el, sourcePath);
          if (headingChildren.length > 0) { docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: headingChildren })); }
          break;
        case 'P':
        case 'DIV':
          if (el.textContent?.trim() || el.querySelector('img')) {
            currentParagraphOptions.spacing = { after: 200 };
            const pChildren = await this.parseInlineElements(el, sourcePath);
            if (pChildren.length > 0) { docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: pChildren })); }
          }
          break;
        case 'UL':
        case 'OL':
          docxObjects.push(...await this.parseListElement(
            el as HTMLUListElement | HTMLOListElement,
            0,
            bodyBgColor,
            sourcePath
          ));
          break;
        case 'HR':
          docxObjects.push(new Paragraph({ thematicBreak: true }));
          docxObjects.push(new Paragraph({ spacing: { after: 200 } }));
          break;
        case 'TABLE':
          docxObjects.push(await this.parseTableElement(el as HTMLTableElement, bodyBgColor, sourcePath));
          docxObjects.push(new Paragraph({}));
          break;
        case 'BLOCKQUOTE':
          const quoteTable = await this.parseQuoteContent(el, sourcePath);
          if (quoteTable) {
            docxObjects.push(quoteTable);
            docxObjects.push(new Paragraph({ spacing: { after: 200 } }));
          }
          break;
        case 'PRE':
          const codeElement = el.querySelector('code');
          if (codeElement) {
            currentParagraphOptions.style = "SourceCode";
            currentParagraphOptions.spacing = {};
            currentParagraphOptions.font = codeFont;
            currentParagraphOptions.shading = { type: ShadingType.CLEAR, fill: 'D3D3D3', color: "auto" };
            const preStyle = window.getComputedStyle(el);
            const codeStyle = window.getComputedStyle(codeElement);
            const langMatch = Array.from(codeElement.classList).find(cls => cls.startsWith('language-'));
            const lang = langMatch ? langMatch.substring('language-'.length) : null;
            const size = this.pxToHalfPoints(codeStyle.fontSize);

            const runs: TextRun[] = [];
            if (lang) {
              runs.push(new TextRun({ text: lang, italics: true, color: "888880", size: (size || 22) - 2 }));
              runs.push(new TextRun({ break: 2 }));
            }
            runs.push(...this.parseSyntaxHighlightedCode(codeElement, codeFont, size));
            docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: runs }));
          }
          break;
        case 'A':
          const linkTextRuns = await this.parseInlineElements(el, sourcePath);
          const hyperlink = new ExternalHyperlink({ link: el.getAttribute('href') || '', children: linkTextRuns });
          docxObjects.push(new Paragraph({ children: [hyperlink], font: mainFont }));
          break;
        default:
          const defaultChildren = await this.parseInlineElements(el, sourcePath);
          if (defaultChildren.length > 0) { docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: defaultChildren, spacing: { after: 200 } })); }
          break;
      }
    }
    return docxObjects;
  }

  private async parseQuoteContent(blockquoteElement: HTMLElement, sourcePath: string, indentLevel: number = 0): Promise<Table | null> {
    const children: (Paragraph | Table)[] = [];
    const nodes = Array.from(blockquoteElement.childNodes);
    const mainFont = { name: 'Times New Roman' };
    const codeFont = { name: 'Courier New' };

    for (const child of nodes) {
      if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        const tagName = el.tagName?.toUpperCase();
        if (!tagName) continue;

        if (tagName === 'BLOCKQUOTE') {
          const nestedQuoteTable = await this.parseQuoteContent(el, sourcePath, indentLevel + 1);
          if (nestedQuoteTable) {
            children.push(nestedQuoteTable);
          }
        } else if (tagName === 'P' || tagName === 'DIV') {
          const inlineChildren = await this.parseInlineElements(el, sourcePath);
          if (inlineChildren.length > 0 || el.textContent?.trim().length > 0) {
            children.push(new Paragraph({
              children: inlineChildren,
              spacing: { after: 100 },
              indent: { left: 200 * (indentLevel + 1) },
              font: mainFont
            }));
          }
        } else if (tagName === 'UL' || tagName === 'OL') {
          const listItems = await this.parseListElementForQuote(el as HTMLUListElement | HTMLOListElement, indentLevel, sourcePath);
          children.push(...listItems);
        } else if (tagName === 'PRE') {
          const codeBlock = await this.parsePreElementForQuote(el, indentLevel, sourcePath);
          if (codeBlock) {
            children.push(codeBlock);
          }
        } else if (tagName === 'TABLE') {
          children.push(new Paragraph({ indent: { left: 200 * (indentLevel + 1) }, font: mainFont }));
          const table = await this.parseTableElement(el as HTMLTableElement, 'F0F0F0', sourcePath);
          children.push(table);
          children.push(new Paragraph({ font: mainFont }));
        } else if (tagName === 'A') {
          const linkTextRuns = await this.parseInlineElements(el, sourcePath);
          const hyperlink = new ExternalHyperlink({ link: el.getAttribute('href') || '', children: linkTextRuns });
          children.push(new Paragraph({ children: [hyperlink], font: mainFont }));
        } else {
          const inlineChildren = await this.parseInlineElements(el, sourcePath);
          if (inlineChildren.length > 0) {
            children.push(new Paragraph({
              children: inlineChildren,
              spacing: { after: 100 },
              indent: { left: 200 * (indentLevel + 1) },
              font: mainFont
            }));
          }
        }
      } else if (child.nodeType === Node.TEXT_NODE && child.textContent) {
        const lines = child.textContent.split('\n');
        for (const line of lines) {
          if (line.trim().length > 0) {
            const tempSpan = document.createElement('span');
            tempSpan.textContent = line.trim();
            const inlineRuns = await this.parseInlineElements(tempSpan, sourcePath);
            children.push(new Paragraph({
              children: inlineRuns,
              spacing: { after: 100 },
              indent: { left: 200 * (indentLevel + 1) },
              font: mainFont
            }));
          }
        }
      }
    }

    if (children.length === 0) {
      return null;
    }

    const quoteCell = new TableCell({
      children: children,
      shading: { type: ShadingType.CLEAR, fill: this.rgbToHex(window.getComputedStyle(blockquoteElement).backgroundColor) || 'F0F0F0', color: "auto" },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        left: { style: BorderStyle.SINGLE, size: 8, color: "auto" },
      },
      margins: {
        left: 200 * indentLevel,
        top: 100,
        bottom: 100,
        right: 100,
      }
    });

    const quoteTable = new Table({
      rows: [
        new TableRow({
          children: [quoteCell],
        }),
      ],
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      width: { size: 100, type: WidthType.PERCENTAGE }
    });

    return quoteTable;
  }

  private async parseListElementForQuote(
    listEl: HTMLUListElement | HTMLOListElement,
    indentLevel: number,
    sourcePath: string,
    numberingRef?: string
  ): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    const listItems = Array.from(listEl.children).filter(child => child.tagName === 'LI');
    const listType = listEl.tagName === 'OL' ? 'number' : 'bullet';
    const mainFont = { name: 'Times New Roman' };

    // 为每个有序列表分配唯一 reference
    let currentNumberingRef = numberingRef;
    if (listType === 'number' && !numberingRef) {
      currentNumberingRef = `default-numbering-${this.numberingCounter++}`;
      this.numberingReferences.add(currentNumberingRef);
    }

    for (const li of listItems) {
      const liElement = li as HTMLLIElement;
      const contentContainer = document.createElement('div');
      let nestedList: HTMLUListElement | HTMLOListElement | null = null;
      for (const child of Array.from(liElement.childNodes)) {
        if (child.nodeType === Node.ELEMENT_NODE && (child.nodeName === 'UL' || child.nodeName === 'OL')) {
          nestedList = child as HTMLUListElement | HTMLOListElement;
        } else {
          contentContainer.appendChild(child.cloneNode(true));
        }
      }

      if (contentContainer.textContent?.trim() || contentContainer.querySelector('img')) {
        const inlineChildren = await this.parseInlineElements(contentContainer, sourcePath);
        const paragraphProperties: any = {
          children: inlineChildren,
          spacing: { after: 100 },
          indent: { left: 720 * (indentLevel + 1) }, // 优化缩进
          font: mainFont
        };

        if (listType === 'bullet') {
          paragraphProperties.bullet = { level: indentLevel };
        } else {
          paragraphProperties.numbering = { reference: currentNumberingRef, level: indentLevel };
        }

        paragraphs.push(new Paragraph(paragraphProperties));
      }

      if (nestedList) {
        paragraphs.push(...await this.parseListElementForQuote(nestedList, indentLevel + 1, sourcePath, currentNumberingRef));
      }
    }
    return paragraphs;
  }

  private async parsePreElementForQuote(preElement: HTMLElement, indentLevel: number, sourcePath: string): Promise<Paragraph | null> {
    const codeElement = preElement.querySelector('code');
    if (!codeElement) return null;

    const preStyle = window.getComputedStyle(preElement);
    const codeStyle = window.getComputedStyle(codeElement);
    const preBgColor = this.rgbToHex(preStyle.backgroundColor) || 'F0F0F0';
    const codeFont = { name: 'Courier New' };
    const size = this.pxToHalfPoints(codeStyle.fontSize);

    const runs: TextRun[] = [];
    const langMatch = Array.from(codeElement.classList).find(cls => cls.startsWith('language-'));
    const lang = langMatch ? langMatch.substring('language-'.length) : null;
    if (lang) {
      runs.push(new TextRun({ text: lang, italics: true, color: "888880", size: (size || 22) - 2 }));
      runs.push(new TextRun({ break: 2 }));
    }
    runs.push(...this.parseSyntaxHighlightedCode(codeElement, codeFont, size));

    const paragraphOptions = {
      style: "SourceCode",
      spacing: { after: 100 },
      indent: { left: 200 * (indentLevel + 1) },
      shading: { type: ShadingType.CLEAR, fill: preBgColor, color: "auto" },
      font: codeFont
    };

    return new Paragraph({ ...paragraphOptions, children: runs });
  }

  // --- 文件保存与主逻辑 ---

  private async saveFile(filePath: string, data: ArrayBuffer) {
    const file = this.app.vault.getAbstractFileByPath(filePath);
    try {
      if (file instanceof TFile) {
        await this.app.vault.modifyBinary(file, data);
      } else {
        await this.app.vault.createBinary(filePath, data);
      }
      new Notice(this.i18n.t("EXPORT_SUCCESSFUL", filePath));
    } catch (error) {
      new Notice(this.i18n.t("SAVE_FAILED", error.message));
    }
  }

  private async fixDocxBlobAuto(blob: Blob): Promise<Blob> {
    const zip = await JSZip.loadAsync(blob);

    const mediaEntries = Object.keys(zip.files).filter(name => name.startsWith('word/media/') && !name.endsWith('/'));
    if (mediaEntries.length === 0) {
      return blob;
    }

    let counter = 0;
    const renameMap: { [key: string]: string } = {};

    for (const oldPath of mediaEntries) {
      const dataU8 = await zip.file(oldPath)!.async('uint8array');
      const detectedMime = this.detectMimeFromHeader(dataU8);
      const ext = this.extFromMime(detectedMime) || (oldPath.match(/\.([a-z0-9]+)$/i) || [null, null])[1] || 'bin';
      counter++;
      const newBase = `image${counter}.${ext}`;
      const newPath = `word/media/${newBase}`;

      const oldLower = oldPath.toLowerCase();
      if (oldLower.endsWith(`.${ext.toLowerCase()}`)) {
        renameMap[oldPath] = oldPath;
        continue;
      }

      zip.file(newPath, dataU8);
      delete zip.files[oldPath];
      renameMap[oldPath] = newPath;
    }

    const relFilePaths = Object.keys(zip.files).filter(n => n.endsWith('.rels'));
    for (const relPath of relFilePaths) {
      let relXml = await zip.file(relPath)!.async('string');
      let changed = false;
      for (const [oldP, newP] of Object.entries(renameMap)) {
        if (oldP === newP) continue;
        const oldNoPrefix = oldP.replace(/^word\//, '');
        const newNoPrefix = newP.replace(/^word\//, '');
        if (relXml.indexOf(oldNoPrefix) !== -1) {
          relXml = relXml.split(oldNoPrefix).join(newNoPrefix);
          changed = true;
        }
        const oldWithWord = oldP;
        const newWithWord = newP;
        if (relXml.indexOf(oldWithWord) !== -1) {
          relXml = relXml.split(oldWithWord).join(newWithWord);
          changed = true;
        }
      }
      if (changed) { zip.file(relPath, relXml); }
    }

    const ctPath = '[Content_Types].xml';
    if (zip.file(ctPath)) {
      let ct = await zip.file(ctPath)!.async('string');
      let ctChanged = false;
      const usedExts = new Set(Object.values(renameMap).map(p => (p.match(/\.([^.]+)$/) || [null, null])[1]).filter(Boolean));
      mediaEntries.forEach(e => { const m = (e.match(/\.([^.]+)$/) || [null, null])[1]; if (m) usedExts.add(m); });

      for (const ext of usedExts) {
        const re = new RegExp(`Extension="${this.escapeRegExp(ext)}"`, 'i');
        if (!re.test(ct)) {
          let contentType = 'image/png';
          if (ext === 'jpg' || ext === 'jpeg') contentType = 'image/jpeg';
          else if (ext === 'gif') contentType = 'image/gif';
          else if (ext === 'bmp') contentType = 'image/bmp';
          else if (ext === 'svg') contentType = 'image/svg+xml';
          else if (ext === 'webp') contentType = 'image/webp';

          ct = ct.replace(/(<Types[^>]*>)/, `$1\n  <Default Extension="${ext}" ContentType="${contentType}"/>`);
          ctChanged = true;
        }
      }
      if (ctChanged) { zip.file(ctPath, ct); }
    }
    return zip.generateAsync({ type: 'blob' });
  }

  async exportCurrentNoteToDocx() {
    const activeFile = this.app.workspace.getActiveFile();
    if (!activeFile) { new Notice(this.i18n.t("NO_ACTIVE_FILE")); return; }

    const tempDiv = document.createElement('div');
    tempDiv.addClass('docx-export-temp-div');

    try {
      document.body.appendChild(tempDiv);
      const markdownContent = await this.app.vault.read(activeFile);
      const sourcePath = activeFile.path;

      const component = new Component();
      await MarkdownRenderer.render(this.app, markdownContent, tempDiv, sourcePath, component);
      component.unload();

      // 重置图片计数器并显示开始导出提示
      this.totalNetworkImages = this.countNetworkImages(tempDiv);
      this.currentImageIndex = 0;
      new Notice(this.i18n.t("EXPORTING_START"));

      const bodyBgColor = this.rgbToHex(window.getComputedStyle(document.body).backgroundColor);

      const docxObjects = await this.htmlToDocxObjects(tempDiv, bodyBgColor, true, 0, sourcePath);

      // 修复：定义 titleParagraph
      const title = activeFile.basename;
      const titleParagraph = new Paragraph({ text: title, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { after: 400 }, font: { name: 'Times New Roman' } });

      // 生成所有需要的 numbering 配置
      const numberingConfig = Array.from(this.numberingReferences).map(ref => ({
        reference: ref,
        levels: [
          { level: 0, format: "decimal", text: "%1.", alignment: AlignmentType.START, indent: { left: 720, hanging: 360 } },
          { level: 1, format: "decimal", text: "%1.%2.", alignment: AlignmentType.START, indent: { left: 1440, hanging: 360 } },
          { level: 2, format: "decimal", text: "%1.%2.%3.", alignment: AlignmentType.START, indent: { left: 2160, hanging: 360 } },
          { level: 3, format: "decimal", text: "%1.%2.%3.%4.", alignment: AlignmentType.START, indent: { left: 2880, hanging: 360 } },
          { level: 4, format: "decimal", text: "%1.%2.%3.%4.%5.", alignment: AlignmentType.START, indent: { left: 3600, hanging: 360 } },
        ],
      }));

      const doc = new Document({
        numbering: {
          config: numberingConfig.length > 0 ? numberingConfig : [{
            reference: "default-numbering",
            levels: [
              { level: 0, format: "decimal", text: "%1.", alignment: AlignmentType.START, indent: { left: 720, hanging: 360 } },
              { level: 1, format: "decimal", text: "%1.%2.", alignment: AlignmentType.START, indent: { left: 1440, hanging: 360 } },
              { level: 2, format: "decimal", text: "%1.%2.%3.", alignment: AlignmentType.START, indent: { left: 2160, hanging: 360 } },
              { level: 3, format: "decimal", text: "%1.%2.%3.%4.", alignment: AlignmentType.START, indent: { left: 2880, hanging: 360 } },
              { level: 4, format: "decimal", text: "%1.%2.%3.%4.%5.", alignment: AlignmentType.START, indent: { left: 3600, hanging: 360 } },
            ],
          }],
        },
        sections: [{
          properties: {},
          children: [titleParagraph, ...docxObjects]
        }]
      });

      const originalBlob = await Packer.toBlob(doc);

      const fixedBlob = await this.fixDocxBlobAuto(originalBlob);

      const buffer = await fixedBlob.arrayBuffer();

      const filePath = activeFile.path.replace(/\.md$/, '.docx');

      const fileExists = await this.app.vault.adapter.exists(filePath);
      if (fileExists) {
        new OverwriteConfirmModal(this.app, async () => {
          await this.saveFile(filePath, buffer);
          new Notice(this.i18n.t("FILE_SAVE_LOCATION_NOTICE"));
        }, this.i18n).open();
      } else {
        await this.saveFile(filePath, buffer);
        new Notice(this.i18n.t("FILE_SAVE_LOCATION_NOTICE"));
      }

    } catch (error) {
      new Notice(this.i18n.t("EXPORT_FAILED"));
    } finally {
      // 重置计数器
      this.totalNetworkImages = 0;
      this.currentImageIndex = 0;
      if (document.body.contains(tempDiv)) {
        document.body.removeChild(tempDiv);
      }
    }
  }
}