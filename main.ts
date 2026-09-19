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
  TableLayoutType,
  ShadingType,
  ExternalHyperlink,
  InternalHyperlink,
  Bookmark,
  PageBreak,
  IBordersOptions,
  BorderStyle,
  ImageRun,
  IParagraphOptions,
  ITableCellOptions
} from 'docx';
import * as JSZip from 'jszip';

// 行内内容：文本、图片、外部链接、指向文档内书签的内部链接，以及书签本身
type InlineRun = TextRun | ExternalHyperlink | InternalHyperlink | ImageRun | Bookmark;

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
    "DOWNLOADING_IMAGE": "Downloading image {0} of {1}...",
    "SVG_CONVERTING": "Converting SVG image: {0}",
    "SVG_CONVERT_FAILED": "SVG conversion failed for: {0}, embedding the original file."
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
    "DOWNLOADING_IMAGE": "正在下载第 {0} 张图片，共 {1} 张...",
    "SVG_CONVERTING": "正在转换 SVG 图片：{0}",
    "SVG_CONVERT_FAILED": "SVG 转换失败：{0}，将嵌入原始文件。"
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
    "DOWNLOADING_IMAGE": "正在下載第 {0} 張圖片，共 {1} 張...",
    "SVG_CONVERTING": "正在轉換 SVG 圖片：{0}",
    "SVG_CONVERT_FAILED": "SVG 轉換失敗：{0}，將嵌入原始檔案。"
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
    "DOWNLOADING_IMAGE": "画像 {0}/{1} をダウンロード中...",
    "SVG_CONVERTING": "SVG画像を変換中: {0}",
    "SVG_CONVERT_FAILED": "SVGの変換に失敗しました: {0}。元のファイルを埋め込みます。"
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
    "DOWNLOADING_IMAGE": "이미지 다운로드 중 ({0}/{1})...",
    "SVG_CONVERTING": "SVG 이미지 변환 중: {0}",
    "SVG_CONVERT_FAILED": "SVG 변환 실패: {0}, 원본 파일을 삽입합니다."
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
    "DOWNLOADING_IMAGE": "Téléchargement de l'image {0} sur {1}...",
    "SVG_CONVERTING": "Conversion de l'image SVG : {0}",
    "SVG_CONVERT_FAILED": "Échec de la conversion SVG pour : {0}, le fichier d'origine sera intégré."
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
    "DOWNLOADING_IMAGE": "Descargando imagen {0} de {1}...",
    "SVG_CONVERTING": "Convirtiendo imagen SVG: {0}",
    "SVG_CONVERT_FAILED": "Error al convertir SVG: {0}, se insertará el archivo original."
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
    "DOWNLOADING_IMAGE": "Загрузка изображения {0} из {1}...",
    "SVG_CONVERTING": "Конвертация SVG-изображения: {0}",
    "SVG_CONVERT_FAILED": "Не удалось конвертировать SVG: {0}, будет вставлен исходный файл."
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
    "DOWNLOADING_IMAGE": "Download immagine {0} di {1}...",
    "SVG_CONVERTING": "Conversione immagine SVG: {0}",
    "SVG_CONVERT_FAILED": "Conversione SVG non riuscita per: {0}, verrà incorporato il file originale."
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
    "DOWNLOADING_IMAGE": "Baixando imagem {0} de {1}...",
    "SVG_CONVERTING": "Convertendo imagem SVG: {0}",
    "SVG_CONVERT_FAILED": "Falha ao converter SVG: {0}, incorporando o arquivo original."
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
    "DOWNLOADING_IMAGE": "Resim indiriliyor {0}/{1}...",
    "SVG_CONVERTING": "SVG görseli dönüştürülüyor: {0}",
    "SVG_CONVERT_FAILED": "SVG dönüştürme başarısız: {0}, orijinal dosya gömülüyor."
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
    "DOWNLOADING_IMAGE": "Lade Bild {0} von {1}...",
    "SVG_CONVERTING": "SVG-Bild wird konvertiert: {0}",
    "SVG_CONVERT_FAILED": "SVG-Konvertierung fehlgeschlagen für: {0}, Originaldatei wird eingebettet."
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
    "DOWNLOADING_IMAGE": "جاري تحميل الصورة {0} من {1}...",
    "SVG_CONVERTING": "جاري تحويل صورة SVG: {0}",
    "SVG_CONVERT_FAILED": "فشل تحويل SVG لـ: {0}، سيتم تضمين الملف الأصلي."
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

  private nodeModuleCache = new Map<string, any>();
  private rsvgConverterPath: string | null | undefined = undefined;
  private svgPngCache = new Map<string, ArrayBuffer>();

  // 标题 -> Word 书签名（用于目录跳转），以及按顺序记录的标题层级信息
  private headingBookmarks = new Map<string, string>();
  private headingBookmarksLoose = new Map<string, string>();
  // 同名标题可能出现多次：每个标题实例都要有唯一书签名，链接指向第一次出现的位置
  private headingBookmarksByKey = new Map<string, string[]>();
  private headingRenderCursor = new Map<string, number>();
  private headingEntries: { level: number, text: string, bookmark: string }[] = [];
  // 原始 Markdown 中 ```table-of-contents 代码块的选项（按出现顺序）
  private tocOptionQueue: string[] = [];

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

  // --- 目录（table-of-contents）与标题书签 ---

  // 标题文本与目录链接目标都用它规范化，保证两边能对上
  private normalizeAnchorText(text: string): string {
    let value = (text ?? '').replace(/\s+/g, ' ').trim();
    // 去掉 Markdown 强调/代码标记（metadataCache 里的标题可能仍带有这些符号）
    value = value.replace(/[*_`~=]/g, '');
    // 目录插件会把 # 与 | 替换成空格，这里保持一致
    value = value.replace(/[#|]/g, ' ');
    return value.replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // 宽松匹配：只保留字母、数字和 CJK，忽略标点与空格差异
  private looseAnchorKey(text: string): string {
    return this.normalizeAnchorText(text).replace(/[^0-9a-z一-鿿]+/g, '');
  }

  // 导出前先扫描所有标题，为它们分配书签（目录通常出现在标题之前，必须提前知道锚点）
  private collectHeadingBookmarks(root: HTMLElement): void {
    this.headingBookmarks.clear();
    this.headingBookmarksLoose.clear();
    this.headingBookmarksByKey.clear();
    this.headingRenderCursor.clear();
    this.headingEntries = [];

    const headings = Array.from(root.querySelectorAll('h1, h2, h3, h4, h5, h6'));
    let counter = 0;
    for (const heading of headings) {
      const element = heading as HTMLElement;
      const text = (element.textContent ?? '').trim();
      const level = Number(element.tagName.substring(1)) || 1;
      if (!text) continue;
      const key = this.normalizeAnchorText(text);
      if (!key) continue;

      counter++;
      // Word 书签名只能由字母、数字、下划线组成，且以字母开头
      const bookmark = `toc_heading_${counter}`;
      this.headingEntries.push({ level, text, bookmark });

      const list = this.headingBookmarksByKey.get(key) ?? [];
      list.push(bookmark);
      this.headingBookmarksByKey.set(key, list);
      if (!this.headingBookmarks.has(key)) this.headingBookmarks.set(key, bookmark);
      const looseKey = this.looseAnchorKey(text);
      if (looseKey && !this.headingBookmarksLoose.has(looseKey)) {
        this.headingBookmarksLoose.set(looseKey, bookmark);
      }
    }
  }

  private getHeadingBookmark(text: string): string | null {
    const key = this.normalizeAnchorText(text);
    if (!key) return null;
    return this.headingBookmarks.get(key) ?? null;
  }

  // 渲染标题时按顺序取用书签名，保证同名标题的书签名也不重复
  private takeHeadingBookmarkForRender(text: string): string | null {
    const key = this.normalizeAnchorText(text);
    if (!key) return null;
    const list = this.headingBookmarksByKey.get(key);
    if (!list || list.length === 0) return null;
    const index = this.headingRenderCursor.get(key) ?? 0;
    this.headingRenderCursor.set(key, index + 1);
    return list[Math.min(index, list.length - 1)];
  }

  // 把 [[#标题]] 这类站内链接解析成书签名
  private resolveHeadingAnchor(rawTarget: string): string | null {
    if (!rawTarget) return null;
    let target = rawTarget.trim();
    if (!target.startsWith('#')) return null;
    target = target.substring(1);
    try { target = decodeURIComponent(target); } catch (error) { }

    // data-href / href 可能形如 "#标题" 或 "#标题|显示文本"
    const candidates = [target, target.split('|')[0], target.replace(/\|/g, ' ')];
    for (const candidate of candidates) {
      const bookmark = this.getHeadingBookmark(candidate);
      if (bookmark) return bookmark;
    }
    const looseKey = this.looseAnchorKey(target);
    return looseKey ? this.headingBookmarksLoose.get(looseKey) ?? null : null;
  }

  // 目录条目文本样式
  private tocLinkTextRun(text: string): TextRun {
    return new TextRun({ text, style: "Hyperlink", color: '0563C1', underline: {} });
  }

  // ```table-of-contents / ```toc 代码块（可能是 pre>code，也可能是已渲染的容器）
  private isTocCodeBlock(el: HTMLElement): boolean {
    const classesOf = (node: HTMLElement) => Array.from(node.classList).map(cls => cls.toLowerCase());
    const classes = classesOf(el);
    const codeEl = el.tagName?.toUpperCase() === 'CODE' ? el : el.querySelector('code');
    const codeClasses = codeEl ? classesOf(codeEl as HTMLElement) : [];
    return [...classes, ...codeClasses].some(cls => cls === 'language-table-of-contents' || cls === 'language-toc');
  }

  // 分页标记：<div class="page-break" style="page-break-before: always;"></div>
  private isPageBreakElement(el: HTMLElement): boolean {
    if (el.classList?.contains('page-break')) return true;
    const styleAttr = (el.getAttribute('style') || '').toLowerCase().replace(/\s+/g, '');
    return /page-break-before:always/.test(styleAttr)
      || /break-before:page/.test(styleAttr)
      || /page-break-after:always/.test(styleAttr);
  }

  // 目录占位元素：可能带 block-language-* 类名，也可能仍是未展开的代码块
  private isTocPlaceholder(el: HTMLElement): boolean {
    if (this.isTocCodeBlock(el)) return true;
    const classes = Array.from(el.classList).map(cls => cls.toLowerCase());
    if (classes.some(cls => cls === 'block-language-table-of-contents' || cls === 'block-language-toc')) return true;
    try {
      if (el.querySelector(':scope > .block-language-table-of-contents, :scope > .block-language-toc')) return true;
    } catch (error) { }
    return false;
  }

  // 从原始 Markdown 中按顺序取出 ```table-of-contents / ```toc 代码块的内容（含选项）
  private parseTocBlocksFromMarkdown(markdown: string): string[] {
    const blocks: string[] = [];
    let inBlock = false;
    let buffer: string[] = [];
    for (const rawLine of (markdown ?? '').split('\n')) {
      const line = rawLine.trim();
      if (!inBlock) {
        if (/^`{3,}\s*(table-of-contents|toc)\s*$/i.test(line)) {
          inBlock = true;
          buffer = [];
        }
      } else if (/^`{3,}\s*$/.test(line)) {
        inBlock = false;
        blocks.push(buffer.join('\n'));
      } else {
        buffer.push(rawLine);
      }
    }
    return blocks;
  }

  // 依次取用 Markdown 里解析到的目录选项，取不到就用传入的文本兜底
  private takeTocOptions(fallback: string): string {
    return this.tocOptionQueue.length > 0 ? this.tocOptionQueue.shift() as string : fallback;
  }

  // 目录代码块所在的容器（Automatic Table Of Contents 插件会生成这类容器）
  private isInsideTocBlock(el: HTMLElement): boolean {
    try {
      return !!el.closest('.block-language-table-of-contents, .block-language-toc');
    } catch (error) {
      return false;
    }
  }

  // 判断列表是否为目录：容器带 TOC 类名，或条目全部是指向本文档标题的链接
  private isTocList(listEl: HTMLElement): boolean {
    if (this.isInsideTocBlock(listEl)) return true;
    const items = Array.from(listEl.querySelectorAll(':scope > li'));
    if (items.length < 2) return false;
    let anchorCount = 0;
    for (const item of items) {
      const link = item.querySelector('a[href^="#"], a[data-href^="#"]');
      if (link && this.resolveHeadingAnchor(link.getAttribute('data-href') || link.getAttribute('href') || '')) {
        anchorCount++;
      }
    }
    return anchorCount >= 2 && anchorCount === items.length;
  }

  // 目录列表渲染为带缩进的段落（不带项目符号，更接近 Word 目录的观感）
  private async parseTocListElement(
    listEl: HTMLUListElement | HTMLOListElement,
    level: number,
    sourcePath: string
  ): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    for (const li of Array.from(listEl.children).filter(c => c.tagName === 'LI')) {
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
      const children = await this.parseInlineElements(contentContainer, sourcePath);
      if (children.length > 0) {
        paragraphs.push(new Paragraph({
          children,
          indent: { left: 360 * level },
          spacing: { after: 60 },
          font: { name: 'Times New Roman' }
        }));
      }
      if (nestedList) {
        paragraphs.push(...await this.parseTocListElement(nestedList, level + 1, sourcePath));
      }
    }
    return paragraphs;
  }

  // 解析 table-of-contents 代码块里的选项（style / minLevel / maxLevel / title 等）
  private parseTocOptions(source: string): { title: string, style: string, minLevel: number, maxLevel: number } {
    const options = { title: '', style: 'nestedList', minLevel: 0, maxLevel: 0 };
    for (const rawLine of (source ?? '').split('\n')) {
      const line = rawLine.trim();
      if (!line || line.startsWith('#')) continue;
      const match = line.match(/^([a-zA-Z]+)\s*:\s*(.*)$/);
      if (!match) continue;
      const name = match[1].toLowerCase();
      // 去掉行尾注释
      let value = match[2].split('#')[0].trim();
      if (name === 'title') {
        options.title = (value === 'null' || value === '""' || value === "''") ? '' : value;
      } else if (name === 'style') {
        options.style = value || 'nestedList';
      } else if (name === 'minlevel' || name === 'maxlevel') {
        const parsed = Number.parseInt(value, 10);
        if (!isNaN(parsed) && parsed >= 0) {
          if (name === 'minlevel') options.minLevel = parsed; else options.maxLevel = parsed;
        }
      }
    }
    return options;
  }

  // 目录插件未启用时，由本插件根据标题生成目录（条目同样是可点击的内部链接）
  private buildTableOfContents(source: string): Paragraph[] {
    if (this.headingEntries.length === 0) return [];
    const options = this.parseTocOptions(source);
    const entries = this.headingEntries.filter(entry =>
      (!options.minLevel || entry.level >= options.minLevel) &&
      (!options.maxLevel || entry.level <= options.maxLevel)
    );
    if (entries.length === 0) return [];

    const baseLevel = options.minLevel > 0
      ? options.minLevel
      : Math.min(...entries.map(entry => entry.level));
    const mainFont = { name: 'Times New Roman' };
    const paragraphs: Paragraph[] = [];

    if (options.title) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: options.title, bold: true, font: mainFont })],
        spacing: { after: 100 },
        font: mainFont
      }));
    }

    if (options.style === 'inlineFirstLevel') {
      const children: InlineRun[] = [];
      entries.filter(entry => entry.level === baseLevel).forEach((entry, index) => {
        if (index > 0) children.push(new TextRun({ text: ' | ', font: mainFont }));
        children.push(new InternalHyperlink({
          anchor: entry.bookmark,
          children: [this.tocLinkTextRun(entry.text)]
        }));
      });
      paragraphs.push(new Paragraph({ children, spacing: { after: 200 }, font: mainFont }));
      return paragraphs;
    }

    for (const entry of entries) {
      const indentLevel = Math.max(0, entry.level - baseLevel);
      paragraphs.push(new Paragraph({
        children: [new InternalHyperlink({
          anchor: entry.bookmark,
          children: [this.tocLinkTextRun(entry.text)]
        })],
        indent: { left: 360 * indentLevel },
        spacing: { after: 60 },
        font: mainFont
      }));
    }
    return paragraphs;
  }

  // --- SVG 转 PNG（调用随插件附带的 rsvg-convert）---

  // 懒加载 Node 模块，移动端不可用
  private requireNodeModule<T = any>(id: string): T | null {
    if (this.nodeModuleCache.has(id)) {
      return (this.nodeModuleCache.get(id) as T) ?? null;
    }
    let mod: any = null;
    try {
      const req: any = (window as any).require
        ?? (globalThis as any).require
        ?? (typeof require === 'function' ? require : null);
      if (req) mod = req(id);
    } catch (error) {
      mod = null;
    }
    this.nodeModuleCache.set(id, mod);
    return mod as T | null;
  }

  // 定位插件目录下的 rsvg-convert 可执行文件
  private resolveRsvgConverterPath(): string | null {
    if (this.rsvgConverterPath !== undefined) return this.rsvgConverterPath;
    let resolved: string | null = null;
    try {
      const pathMod = this.requireNodeModule<any>('path');
      const fsMod = this.requireNodeModule<any>('fs');
      const adapter: any = (this.app.vault as any).adapter;
      const basePath: string = typeof adapter?.getBasePath === 'function' ? adapter.getBasePath() : '';
      if (pathMod && fsMod && basePath) {
        const hasProcess = typeof process !== 'undefined';
        const isWindows = hasProcess && process.platform === 'win32';
        const fileNames = isWindows ? ['rsvg-convert.exe', 'rsvg-convert'] : ['rsvg-convert', 'rsvg-convert.exe'];
        // 优先使用插件自带的二进制，其次回退到系统 PATH 中安装的 rsvg-convert
        const pathEnv: string = hasProcess ? (process.env?.PATH ?? '') : '';
        const dirs = [
          this.manifest.dir,
          `.obsidian/plugins/${this.manifest.id}`,
          ...pathEnv.split(pathMod.delimiter).filter(Boolean)
        ];
        for (const dir of dirs) {
          if (!dir) continue;
          for (const fileName of fileNames) {
            const candidate = pathMod.join(basePath, dir, fileName);
            try {
              if (fsMod.existsSync(candidate)) { resolved = candidate; break; }
            } catch (error) { }
          }
          if (resolved) break;
        }
      }
    } catch (error) {
      resolved = null;
    }
    this.rsvgConverterPath = resolved;
    return resolved;
  }

  // 判断图片数据是否为 SVG
  private looksLikeSvg(buffer: ArrayBuffer, extension: string | null): boolean {
    if (extension && extension.toLowerCase().replace(/^\./, '') === 'svg') return true;
    try {
      const head = new Uint8Array(buffer.slice(0, 512));
      return new TextDecoder().decode(head).includes('<svg');
    } catch (error) {
      return false;
    }
  }

  // 从 width/height/viewBox 解析 SVG 固有尺寸
  private getSvgIntrinsicSize(svgText: string): { width: number, height: number } | null {
    const header = svgText.slice(0, 4096);
    const readLength = (name: string): { value: number, unit: string } | null => {
      const match = header.match(new RegExp(`\\s${name}\\s*=\\s*["']([^"']+)["']`, 'i'));
      if (!match) return null;
      const value = parseFloat(match[1]);
      if (isNaN(value) || value <= 0) return null;
      const unit = (match[1].match(/[a-z%]+/i)?.[0] || 'px').toLowerCase();
      return { value, unit };
    };
    const width = readLength('width');
    const height = readLength('height');
    if (width && height && width.unit !== '%' && height.unit !== '%') {
      return {
        width: this.svgLengthToPx(width.value, width.unit),
        height: this.svgLengthToPx(height.value, height.unit)
      };
    }
    const viewBox = header.match(/viewBox\s*=\s*["']([^"']+)["']/i);
    if (viewBox) {
      const parts = viewBox[1].trim().split(/[\s,]+/).map(Number);
      if (parts.length === 4 && parts[2] > 0 && parts[3] > 0) {
        return { width: parts[2], height: parts[3] };
      }
    }
    return null;
  }

  // SVG 长度单位转 px
  private svgLengthToPx(value: number, unit: string): number {
    switch (unit) {
      case 'pt': return value * 96 / 72;
      case 'pc': return value * 16;
      case 'in': return value * 96;
      case 'cm': return value * 96 / 2.54;
      case 'mm': return value * 96 / 25.4;
      case 'em': case 'ex': case 'rem': return value * 16;
      default: return value;
    }
  }

  // 计算 SVG 在文档中应显示的尺寸
  private resolveSvgDisplaySize(
    imgEl: HTMLImageElement,
    intrinsic: { width: number, height: number } | null
  ): { width: number, height: number } {
    const styleWidth = parseFloat(imgEl.style.width) || imgEl.width || 0;
    const styleHeight = parseFloat(imgEl.style.height) || imgEl.height || 0;
    const naturalWidth = imgEl.naturalWidth > 0 ? imgEl.naturalWidth : 0;
    const naturalHeight = imgEl.naturalHeight > 0 ? imgEl.naturalHeight : 0;

    let width = styleWidth || intrinsic?.width || naturalWidth || 0;
    let height = styleHeight || intrinsic?.height || naturalHeight || 0;

    if (width > 0 && height <= 0) {
      const ratio = intrinsic
        ? intrinsic.height / intrinsic.width
        : (naturalWidth && naturalHeight ? naturalHeight / naturalWidth : 0.75);
      height = width * ratio;
    } else if (height > 0 && width <= 0) {
      const ratio = intrinsic
        ? intrinsic.width / intrinsic.height
        : (naturalWidth && naturalHeight ? naturalWidth / naturalHeight : 4 / 3);
      width = height * ratio;
    }

    if (width <= 0 || height <= 0) { width = 550; height = 300; }
    return { width: Math.round(width), height: Math.round(height) };
  }

  private hashText(text: string): string {
    let hash = 5381;
    for (let i = 0; i < text.length; i++) {
      hash = ((hash << 5) + hash + text.charCodeAt(i)) >>> 0;
    }
    return hash.toString(36);
  }

  // 调用 rsvg-convert 将 SVG 渲染为 PNG（按显示宽度的 2 倍渲染以保证清晰度）
  private async convertSvgToPng(svgText: string, displayWidth: number): Promise<ArrayBuffer | null> {
    const cacheKey = `${this.hashText(svgText)}-${displayWidth}`;
    const cached = this.svgPngCache.get(cacheKey);
    if (cached) return cached;

    const fsMod = this.requireNodeModule<any>('fs');
    const osMod = this.requireNodeModule<any>('os');
    const pathMod = this.requireNodeModule<any>('path');
    const childProcessMod = this.requireNodeModule<any>('child_process');
    const converterPath = this.resolveRsvgConverterPath();
    if (!fsMod || !osMod || !pathMod || !childProcessMod?.execFile || !converterPath) return null;

    const tmpDir = osMod.tmpdir();
    const uniqueId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    const svgFilePath = pathMod.join(tmpDir, `obsidian-docx-${uniqueId}.svg`);
    const pngFilePath = pathMod.join(tmpDir, `obsidian-docx-${uniqueId}.png`);
    const renderWidth = Math.max(32, Math.min(3000, Math.round(displayWidth * 2)));

    const cleanup = () => {
      try { if (fsMod.existsSync(svgFilePath)) fsMod.unlinkSync(svgFilePath); } catch (error) { }
      try { if (fsMod.existsSync(pngFilePath)) fsMod.unlinkSync(pngFilePath); } catch (error) { }
    };

    try {
      await fsMod.promises.writeFile(svgFilePath, new TextEncoder().encode(svgText));
      const argSets: string[][] = [
        ['-f', 'png', '-a', '-b', 'white', '-w', String(renderWidth), '-o', pngFilePath, svgFilePath],
        ['-f', 'png', '-a', '-w', String(renderWidth), '-o', pngFilePath, svgFilePath]
      ];
      let lastError: any = null;
      for (const args of argSets) {
        lastError = await new Promise<any>(resolve => {
          childProcessMod.execFile(
            converterPath,
            args,
            { windowsHide: true, timeout: 30000, cwd: tmpDir },
            (error: any) => resolve(error ?? null)
          );
        });
        if (!lastError) break;
      }
      if (lastError) {
        cleanup();
        return null;
      }

      const data: Uint8Array = await fsMod.promises.readFile(pngFilePath);
      const result = new Uint8Array(data.length);
      result.set(data);
      this.svgPngCache.set(cacheKey, result.buffer);
      cleanup();
      return result.buffer;
    } catch (error) {
      cleanup();
      return null;
    }
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

  // docx 默认页面（A4 11906 twips + 左右各 1 英寸页边距 1440）下正文可用宽度
  private readonly pageContentTwips = 9026;
  private readonly pageContentWidthPx = 602;
  // 单元格左右内边距：各 100 twips
  private readonly cellMarginPx = 2 * (100 / 20) * (96 / 72);

  // 表格列数（考虑 colspan，取各行的最大值）
  private countTableColumns(tableEl: HTMLElement): number {
    let max = 0;
    for (const row of Array.from(tableEl.querySelectorAll('tr'))) {
      let count = 0;
      for (const cell of Array.from(row.children)) {
        const tag = (cell as HTMLElement).tagName?.toUpperCase();
        if (tag !== 'TD' && tag !== 'TH') continue;
        count += Number(cell.getAttribute('colspan') || '1') || 1;
      }
      max = Math.max(max, count);
    }
    return max;
  }

  // 各列宽度占比：优先用渲染后的真实宽度（取第一行），量不到时按列数均分
  private computeColumnFractions(tableEl: HTMLElement, columns: number): number[] {
    const count = Math.max(1, columns);
    const widths: number[] = new Array(count).fill(0);
    let measured = false;
    for (const row of Array.from(tableEl.querySelectorAll('tr'))) {
      const cells = Array.from(row.children).filter(c => ['TD', 'TH'].includes((c as HTMLElement).tagName));
      cells.forEach((cell, index) => {
        if (index >= count) return;
        const width = (cell as HTMLElement).getBoundingClientRect().width;
        if (width > 0) { measured = true; widths[index] = Math.max(widths[index], width); }
      });
      if (measured) break;
    }
    if (!measured) return new Array(count).fill(1 / count);
    const total = widths.reduce((sum, w) => sum + w, 0);
    if (total <= 0) return new Array(count).fill(1 / count);
    return widths.map(w => w / total);
  }

  // 图片位于表格单元格时，返回该单元格可容纳的宽度（px），否则返回 null
  // 与 parseTableElement 里设置的列宽使用同一套列宽占比，保证图片不会把单元格撑开
  private getTableCellAvailableWidth(imgEl: HTMLElement): number | null {
    let cell: HTMLElement | null = null;
    try { cell = imgEl.closest('td, th') as HTMLElement | null; } catch (error) { return null; }
    if (!cell) return null;
    const table = cell.closest('table') as HTMLElement | null;
    if (!table) return null;

    const columns = this.countTableColumns(table);
    const fractions = this.computeColumnFractions(table, columns);
    const row = cell.parentElement as HTMLElement | null;
    let index = 0;
    if (row) {
      const siblings = Array.from(row.children).filter(c => ['TD', 'TH'].includes((c as HTMLElement).tagName));
      const found = siblings.indexOf(cell);
      if (found >= 0) index = found;
    }
    const fraction = fractions[index] ?? (1 / Math.max(1, columns));
    // 扣掉单元格左右边距与边框，再留一点余量
    const available = this.pageContentWidthPx * fraction - this.cellMarginPx - 6;
    return Math.max(40, Math.round(available));
  }

  // 按最大宽度等比缩放
  private clampImageSize(size: { width: number, height: number }, maxWidth: number): { width: number, height: number } {
    if (!size.width || size.width <= maxWidth) return size;
    const height = Math.max(1, Math.round((maxWidth / size.width) * size.height));
    return { width: Math.round(maxWidth), height };
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
    const columnCount = this.countTableColumns(tableEl);
    const columnFractions = this.computeColumnFractions(tableEl, columnCount);

    // 遍历表格行
    for (const row of Array.from(tableEl.querySelectorAll('tr'))) {
      const cells: TableCell[] = [];
      let cellIndex = 0;

      // 遍历单元格
      for (const cell of Array.from(row.querySelectorAll('th, td'))) {
        const cellStyle = window.getComputedStyle(cell);
        // 用 dxa（缇）指定列宽：docx 的百分比宽度会写成非法的 w:w="100%"，Word 只能退回自动调整
        const columnTwips = Math.max(
          1,
          Math.round(this.pageContentTwips * (columnFractions[cellIndex] ?? (1 / Math.max(1, columnCount))))
        );
        cellIndex++;
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
          width: { size: columnTwips, type: WidthType.DXA },
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
      width: { size: this.pageContentTwips, type: WidthType.DXA },
      layout: TableLayoutType.FIXED,
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
  private async parseInlineElements(element: HTMLElement, sourcePath: string): Promise<InlineRun[]> {
    const runs: InlineRun[] = [];
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
            const linkText = el.textContent?.trim() || '';
            const linkChildren: InlineRun[] = [new TextRun({ text: linkText, style: "Hyperlink", color: this.rgbToHex(style.color) || '0563C1', underline: {} })];
            const anchor = this.resolveHeadingAnchor(el.getAttribute('data-href') || el.getAttribute('href') || '');
            if (anchor) {
              // 指向本文档标题（目录条目）：用书签实现可点击跳转
              runs.push(new InternalHyperlink({ anchor, children: linkChildren }));
            } else if ((el.getAttribute('href') || '').startsWith('#')) {
              // 站内锚点但没匹配到标题，退化为普通文本，避免生成无效链接
              runs.push(new TextRun({ text: linkText, color: this.rgbToHex(style.color) || '0563C1', underline: {} }));
            } else {
              runs.push(new ExternalHyperlink({ link: el.getAttribute('href') || '', children: linkChildren }));
            }
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

  // vault 在磁盘上的根路径（仅桌面端可用）
  private getVaultBasePath(): string {
    try {
      const adapter: any = (this.app.vault as any).adapter;
      return typeof adapter?.getBasePath === 'function' ? adapter.getBasePath() : '';
    } catch (error) {
      return '';
    }
  }

  // 按库内路径 / 链接名查找文件
  private findVaultFile(linkPath: string, sourcePath: string): TFile | null {
    if (!linkPath) return null;
    const normalized = linkPath.replace(/\\/g, '/').replace(/^\/+/, '');
    const abstract = this.app.vault.getAbstractFileByPath(normalized);
    if (abstract instanceof TFile) return abstract;
    const dest = this.app.metadataCache.getFirstLinkpathDest(normalized, sourcePath);
    if (dest instanceof TFile) return dest;
    // 兜底：忽略大小写再匹配一次完整路径
    const lower = normalized.toLowerCase();
    const matched = this.app.vault.getFiles().find(file => file.path.toLowerCase() === lower);
    return matched ?? null;
  }

  // 从 img 的 src（app://local/...、capacitor://... 或库内相对路径）反查库内文件
  // 标准 Markdown 语法 ![](path) 不会被包裹成 internal-embed，只能靠 src 解析
  private resolveImageFileFromSrc(src: string, sourcePath: string): TFile | null {
    const withoutScheme = src.replace(/^[a-z][a-z0-9+.-]*:\/\/[^/]*\/?/i, '');
    const queryPath = src.match(/[?&]path=([^&]+)/i)?.[1];
    const rawCandidates = [withoutScheme.split(/[?#]/)[0], queryPath].filter(Boolean) as string[];

    const basePath = this.getVaultBasePath().replace(/\\/g, '/').replace(/\/+$/, '');
    const candidates: string[] = [];
    for (const rawCandidate of rawCandidates) {
      let decoded = rawCandidate;
      try { decoded = decodeURIComponent(rawCandidate); } catch (error) { }
      decoded = decoded.replace(/\\/g, '/').replace(/^\/+/, '');
      if (!decoded) continue;
      if (basePath && decoded.toLowerCase().startsWith(`${basePath.toLowerCase()}/`)) {
        candidates.push(decoded.substring(basePath.length + 1));
      }
      candidates.push(decoded);
    }

    for (const candidate of candidates) {
      const file = this.findVaultFile(candidate, sourcePath);
      if (file) return file;
    }
    // 绝对路径无法对应到库内时，退化为按文件名查找
    for (const candidate of candidates) {
      const fileName = candidate.split('/').pop();
      if (fileName) {
        const file = this.findVaultFile(fileName, sourcePath);
        if (file) return file;
      }
    }
    return null;
  }

  private async createImageRun(imgEl: HTMLImageElement, sourcePath: string): Promise<ImageRun | null> {
    const src = imgEl.getAttribute('src') ?? '';
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
        const pathFromEmbed = parentSpan?.getAttribute('alt') || parentSpan?.getAttribute('data-href') || parentSpan?.getAttribute('data-src');
        let imageFile: TFile | null = null;

        if (pathFromEmbed) {
          pathForNotice = pathFromEmbed;
          imageFile = this.findVaultFile(pathFromEmbed, sourcePath);
          if (!imageFile) {
            new Notice(this.i18n.t("LOCAL_IMAGE_NOT_FOUND", pathFromEmbed));
            return null;
          }
        } else {
          // 标准 Markdown 图片语法（如 ![](img.svg)、表格里的 ![](img.png)）没有 internal-embed 包裹，
          // 只能从 src 反查库内文件
          imageFile = this.resolveImageFileFromSrc(src, sourcePath);
          if (!imageFile) {
            console.error('[DOCX Exporter] Cannot resolve local image from src:', src);
            new Notice(this.i18n.t("IMAGE_LINK_MISSING"));
            return null;
          }
          pathForNotice = imageFile.path;
        }

        try {
          buffer = await this.app.vault.readBinary(imageFile);
          imageExtension = imageFile.extension;
        } catch (readError) {
          new Notice(this.i18n.t("FILE_READ_FAILED", imageFile.path));
          return null;
        }

      } else if (src.startsWith('data:image')) {
        pathForNotice = 'Base64 embedded image';
        const base64String = src.split(',')[1];
        buffer = this.base64ToArrayBuffer(base64String);
        const mimeType = src.match(/data:image\/(.*?);/)?.[1];
        imageExtension = mimeType || 'png';
      } else {
        // 未带协议前缀的相对路径（部分渲染场景），先尝试按库内文件解析
        const localFile = this.resolveImageFileFromSrc(src, sourcePath);
        if (!localFile) {
          new Notice(this.i18n.t("UNSUPPORTED_IMAGE_FORMAT", src));
          return null;
        }
        pathForNotice = localFile.path;
        try {
          buffer = await this.app.vault.readBinary(localFile);
          imageExtension = localFile.extension;
        } catch (readError) {
          new Notice(this.i18n.t("FILE_READ_FAILED", localFile.path));
          return null;
        }
      }

      if (!buffer || !imageExtension) {
        new Notice(this.i18n.t("IMAGE_DATA_INVALID"));
        return null;
      }

      // 表格单元格里的图片必须收敛到单元格宽度内，否则 Word 中会显示不全
      const maxImageWidth = this.getTableCellAvailableWidth(imgEl) ?? 550;

      // SVG 无法被 Word 稳定渲染，先用 rsvg-convert 栅格化为 PNG 再嵌入
      let svgDisplaySize: { width: number, height: number } | null = null;
      if (this.looksLikeSvg(buffer, imageExtension)) {
        const svgText = new TextDecoder().decode(new Uint8Array(buffer));
        svgDisplaySize = this.resolveSvgDisplaySize(imgEl, this.getSvgIntrinsicSize(svgText));
        // 先按可用宽度收敛，再按收敛后的宽度渲染 PNG，避免生成超大位图
        svgDisplaySize = this.clampImageSize(svgDisplaySize, maxImageWidth);
        // 只有在 rsvg-convert 可用时才转换（移动端没有 Node 环境，保持原样嵌入）
        if (this.resolveRsvgConverterPath()) {
          const shortPath = pathForNotice.length > 50 ? `${pathForNotice.substring(0, 50)}...` : pathForNotice;
          new Notice(this.i18n.t("SVG_CONVERTING", shortPath));
          const pngBuffer = await this.convertSvgToPng(svgText, svgDisplaySize.width);
          if (pngBuffer) {
            buffer = pngBuffer;
            imageExtension = 'png';
          } else {
            new Notice(this.i18n.t("SVG_CONVERT_FAILED", shortPath));
          }
        }
      }

      const maxWidth = maxImageWidth;
      let finalWidth: number;
      let finalHeight: number;

      if (svgDisplaySize) {
        finalWidth = svgDisplaySize.width;
        finalHeight = svgDisplaySize.height;
      } else {
        const dimensionsFromBuffer = this.getImageDimensionsFromBuffer(buffer);
        let naturalWidth = dimensionsFromBuffer?.width;
        let naturalHeight = dimensionsFromBuffer?.height;

        if (!naturalWidth || !naturalHeight) {
          naturalWidth = imgEl.naturalWidth > 0 ? imgEl.naturalWidth : undefined;
          naturalHeight = imgEl.naturalHeight > 0 ? imgEl.naturalHeight : undefined;
        }

        const styleWidth = parseFloat(imgEl.style.width) || imgEl.width;
        const styleHeight = parseFloat(imgEl.style.height) || imgEl.height;

        finalWidth = styleWidth || naturalWidth || 0;
        finalHeight = styleHeight || (finalWidth && naturalWidth && naturalHeight ? (finalWidth / naturalWidth) * naturalHeight : 0);

        if (!finalWidth || !finalHeight || finalWidth <= 0 || finalHeight <= 0) {
          finalWidth = maxWidth;
          finalHeight = 300;
        } else {
          finalWidth = Math.round(finalWidth);
          finalHeight = Math.round(finalHeight);
        }
      }

      finalWidth = Math.round(finalWidth);
      finalHeight = Math.round(finalHeight);
      if (finalWidth <= 0) finalWidth = maxWidth;
      if (finalHeight <= 0) finalHeight = 300;
      if (finalWidth > maxWidth) {
        finalHeight = Math.round((maxWidth / finalWidth) * finalHeight);
        finalWidth = maxWidth;
      }

      // 便于排查：控制台里能看到每张图片最终用的宽度上限与尺寸
      console.log(`[DOCX Exporter] image: ${pathForNotice} | maxWidth=${maxImageWidth} | size=${finalWidth}x${finalHeight}`);

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

      // 分页标记：插入 Word 分页符
      if (this.isPageBreakElement(el)) {
        docxObjects.push(new Paragraph({ children: [new PageBreak()] }));
        continue;
      }

      // 目录占位块（```table-of-contents / ```toc）：直接用标题生成可跳转目录，
      // 不依赖 Automatic Table Of Contents 插件渲染出来的 DOM
      if (this.isTocPlaceholder(el)) {
        docxObjects.push(...this.buildTableOfContents(this.takeTocOptions(el.textContent ?? '')));
        continue;
      }

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
          if (headingChildren.length > 0) {
            // 给标题加书签，目录里的内部链接才能跳转过来
            const bookmark = this.takeHeadingBookmarkForRender(el.textContent ?? '');
            const headingRuns: InlineRun[] = bookmark
              ? [new Bookmark({ id: bookmark, children: headingChildren })]
              : headingChildren;
            docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: headingRuns }));
          }
          break;
        case 'P':
        case 'DIV':
          // 容器里直接包含列表时按块级结构递归解析，避免列表被压成一整段纯文本
          if (el.querySelector(':scope > ul, :scope > ol')) {
            docxObjects.push(...await this.htmlToDocxObjects(el, bodyBgColor, false, indentLevel, sourcePath));
            break;
          }
          // 分页标记可能被包在其它容器里，递归下去才能识别
          if (el.querySelector('.page-break')) {
            docxObjects.push(...await this.htmlToDocxObjects(el, bodyBgColor, false, indentLevel, sourcePath));
            break;
          }
          if (el.textContent?.trim() || el.querySelector('img')) {
            currentParagraphOptions.spacing = { after: 200 };
            const pChildren = await this.parseInlineElements(el, sourcePath);
            if (pChildren.length > 0) { docxObjects.push(new Paragraph({ ...currentParagraphOptions, children: pChildren })); }
          }
          break;
        case 'UL':
        case 'OL':
          const listElement = el as HTMLUListElement | HTMLOListElement;
          if (this.isTocList(listElement)) {
            docxObjects.push(...await this.parseTocListElement(listElement, 0, sourcePath));
          } else {
            docxObjects.push(...await this.parseListElement(listElement, 0, bodyBgColor, sourcePath));
          }
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
          if (codeElement && this.isTocCodeBlock(codeElement)) {
            // 目录插件未启用时，代码块不会被展开，这里自行生成可跳转的目录
            docxObjects.push(...this.buildTableOfContents(this.takeTocOptions(el.textContent ?? '')));
            break;
          }
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
          docxObjects.push(new Paragraph({ children: linkTextRuns, font: mainFont }));
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
          children.push(new Paragraph({ children: linkTextRuns, font: mainFont }));
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

  // docx 库给每个 Bookmark 都生成同样的数字 id，Word 会因为 id 重复而报“内容有问题”，
  // 这里按顺序重新编号，保证 w:id 唯一
  private async fixBookmarkIdsInZip(zip: JSZip): Promise<void> {
    const docPath = 'word/document.xml';
    const docFile = zip.file(docPath);
    if (!docFile) return;

    let xml = await docFile.async('string');
    if (!xml.includes('w:bookmarkStart')) return;

    let startCounter = 0;
    xml = xml.replace(/<w:bookmarkStart\b([^>]*?)\/?>/g, (_match, attrs: string) => {
      startCounter++;
      return `<w:bookmarkStart${attrs.replace(/w:id="[^"]*"/, `w:id="${startCounter}"`)}/>`;
    });
    let endCounter = 0;
    xml = xml.replace(/<w:bookmarkEnd\b([^>]*?)\/?>/g, (_match, attrs: string) => {
      endCounter++;
      return `<w:bookmarkEnd${attrs.replace(/w:id="[^"]*"/, `w:id="${endCounter}"`)}/>`;
    });
    zip.file(docPath, xml);
  }

  // 固定布局下 Word 以 w:tblGrid 的列宽为准，而 docx 生成的是写死的 100 缇，
  // 这里用单元格实际宽度重写列宽，表格才不会超出页面、图片才不会撑破单元格
  private async fixTableGridInZip(zip: JSZip): Promise<void> {
    const docPath = 'word/document.xml';
    const docFile = zip.file(docPath);
    if (!docFile) return;

    let xml = await docFile.async('string');
    if (!xml.includes('<w:tblGrid>')) return;

    xml = xml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, (tableXml: string) => {
      const firstRow = tableXml.match(/<w:tr>[\s\S]*?<\/w:tr>/);
      const widths = firstRow
        ? [...firstRow[0].matchAll(/<w:tcW[^>]*w:w="(\d+)"[^>]*\/>/g)]
          .map(match => Number(match[1]))
          .filter(value => value > 0)
        : [];

      if (widths.length === 0) {
        // 取不到列宽就退回自动布局，避免列被压成 100 缇
        return tableXml.replace(/<w:tblLayout[^>]*\/>/g, '');
      }

      const grid = `<w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`;
      let patched = tableXml.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, grid);
      if (!patched.includes('<w:tblLayout')) {
        patched = patched.replace(/<\/w:tblPr>/, '<w:tblLayout w:type="fixed"/></w:tblPr>');
      }
      return patched;
    });

    zip.file(docPath, xml);
  }

  private async fixDocxBlobAuto(blob: Blob): Promise<Blob> {
    const zip = await JSZip.loadAsync(blob);

    // 修正书签 id（目录跳转依赖它）
    await this.fixBookmarkIdsInZip(zip);

    // 修正表格列宽（固定布局依赖 tblGrid）
    await this.fixTableGridInZip(zip);

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

      // 先扫描全部标题并建立书签映射，保证目录链接（通常出现在标题之前）能正确指向
      this.collectHeadingBookmarks(tempDiv);
      // 从原始 Markdown 里取出目录代码块的选项
      this.tocOptionQueue = this.parseTocBlocksFromMarkdown(markdownContent);

      // 重置图片计数器并显示开始导出提示
      this.totalNetworkImages = this.countNetworkImages(tempDiv);
      this.currentImageIndex = 0;
      new Notice(this.i18n.t("EXPORTING_START"));

      const bodyBgColor = this.rgbToHex(window.getComputedStyle(document.body).backgroundColor);

      const docxObjects = await this.htmlToDocxObjects(tempDiv, bodyBgColor, true, 0, sourcePath);

      // 标题完全沿用 Markdown 的层级：# -> 一级标题，## -> 二级标题……不再插入文件名标题
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
          children: docxObjects
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