import {
  App,
  Plugin,
  Notice,
  MarkdownRenderer,
  Modal,
  TFile,
  requestUrl
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
    "INVALID_FILE_PATH": "Invalid file path. Please enter a valid path."
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
    "FILE_SAVE_LOCATION_NOTICE": "已将 DOCX 文件导出到和笔记文件相同的文件夹内。"
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
    "FILE_SAVE_LOCATION_NOTICE": "已將 DOCX 文件匯出到和筆記文件相同的資料夾內。"
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
    "FILE_SAVE_LOCATION_NOTICE": "DOCXファイルはノートと同じフォルダにエクスポートされました。"
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
    "EXPORT_FAILED": "DOCX 내보내기 실패. 자세한 내용은 개발자 콘솔을 확인하십시오.",
    "SAVE_FAILED": "파일 저장 실패: {0}",
    "DOWNLOAD_IMAGE_FAILED": "이미지 다운로드 실패: {0}",
    "IMAGE_LINK_MISSING": "이미지 링크 속성이 누락되었습니다. 내보내기를 건너뜁니다.",
    "LOCAL_IMAGE_NOT_FOUND": "로컬 이미지를 찾을 수 없습니다: {0}. 내보내기를 건너뜁니다.",
    "LOCAL_IMAGE_FILE_EMPTY": "로컬 이미지 파일이 비어 있습니다. 내보내기를 건너뜁니다.",
    "UNSUPPORTED_IMAGE_FORMAT": "지원되지 않는 이미지 형식: {0}",
    "IMAGE_DATA_INVALID": "이미지 데이터가 잘못되었습니다. 내보내기를 건너뜁니다.",
    "IMAGE_PROCESSING_ERROR": "이미지 처리 중 오류가 발생했습니다: {0}",
    "FILE_READ_FAILED": "로컬 이미지 파일을 읽는 데 실패했습니다: {0}. 파일 권한을 확인하십시오.",
    "SAVE_FILE_DIALOG_TITLE": "DOCX 파일 저장",
    "SAVE_FILE_INPUT_LABEL": "파일 경로 (볼트 루트 기준):",
    "BUTTON_SAVE": "저장",
    "INVALID_FILE_PATH": "잘못된 파일 경로입니다. 유효한 경로를 입력하십시오.",
    "FILE_SAVE_LOCATION_NOTICE": "DOCX 파일은 노트와 동일한 폴더에 내보내졌습니다."
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
    "FILE_SAVE_LOCATION_NOTICE": "Le fichier DOCX a été exporté dans le même dossier que le fichier de note."
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
    "FILE_SAVE_LOCATION_NOTICE": "El archivo DOCX ha sido exportado a la misma carpeta que el archivo de notas."
  },
  pt: {
    "PLUGIN_NAME": "Exportador DOCX",
    "EXPORT_COMMAND_NAME": "Exportar nota atual para DOCX",
    "LOADING_PLUGIN": "Carregando plugin de exportação DOCX",
    "UNLOADING_PLUGIN": "Descarregando plugin de exportação DOCX",
    "NO_ACTIVE_FILE": "Nenhum arquivo ativo para exportar.",
    "FILE_ALREADY_EXISTS_TITLE": "O arquivo já existe",
    "OVERWRITE_CONFIRMATION": "Deseja substituir o arquivo .docx existente?",
    "BUTTON_OVERWRITE": "Substituir",
    "BUTTON_CANCEL": "Cancelar",
    "EXPORT_SUCCESSFUL": "Exportado com sucesso para \"{0}\"",
    "EXPORT_FAILED": "Falha ao exportar DOCX. Verifique o console do desenvolvedor para obter detalhes.",
    "SAVE_FAILED": "Falha ao salvar o arquivo: {0}",
    "DOWNLOAD_IMAGE_FAILED": "Falha ao baixar a imagem: {0}",
    "IMAGE_LINK_MISSING": "O atributo de link da imagem está ausente, ignorando a exportação.",
    "LOCAL_IMAGE_NOT_FOUND": "Imagem local não encontrada: {0}, ignorando a exportação.",
    "LOCAL_IMAGE_FILE_EMPTY": "O arquivo de imagem local está vazio, ignorando a exportação.",
    "UNSUPPORTED_IMAGE_FORMAT": "Formato de imagem não suportado para: {0}",
    "IMAGE_DATA_INVALID": "Dados de imagem inválidos, ignorando a exportação.",
    "IMAGE_PROCESSING_ERROR": "Erro ao processar a imagem: {0}",
    "FILE_READ_FAILED": "Falha ao ler o arquivo de imagem local: {0}, verifique as permissões do arquivo.",
    "SAVE_FILE_DIALOG_TITLE": "Salvar arquivo DOCX",
    "SAVE_FILE_INPUT_LABEL": "Caminho do arquivo (relativo à raiz do cofre):",
    "BUTTON_SAVE": "Salvar",
    "INVALID_FILE_PATH": "Caminho do arquivo inválido. Por favor, insira um caminho válido.",
    "FILE_SAVE_LOCATION_NOTICE": "O arquivo DOCX foi exportado para a mesma pasta que o arquivo de nota."
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
    "FILE_SAVE_LOCATION_NOTICE": "Файл DOCX был экспортирован в ту же папку, что и файл заметки."
  }
};

class I18N {
  private lang: string;
  private translations: Record<string, string>;

  constructor(app: App) {
    let lang = app.vault.config.language;
    if (!lang) {
        lang = window.localStorage.getItem('language') || 'en';
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

  async onload() {
    this.i18n = new I18N(this.app);
    console.log(this.i18n.t("LOADING_PLUGIN"));
    
    this.addRibbonIcon('file-output', this.i18n.t("EXPORT_COMMAND_NAME"), () => this.exportCurrentNoteToDocx());
    this.addCommand({
      id: 'export-to-docx',
      name: this.i18n.t("EXPORT_COMMAND_NAME"),
      callback: () => this.exportCurrentNoteToDocx()
    });
  }
  
  onunload() {
    console.log(this.i18n.t("UNLOADING_PLUGIN"));
  }

  // --- 辅助函数 ---
  private pxToHalfPoints(px: string): number | undefined {
    const pxValue = parseFloat(px);
    if (isNaN(pxValue)) return undefined;
    return Math.round(pxValue * 0.75 * 2);
  }

  private rgbToHex(rgb: string): string | undefined {
    if (!rgb || !rgb.startsWith('rgb') || rgb === 'rgba(0, 0, 0, 0)') return undefined;
    const parts = rgb.match(/^rgb(?:a)?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d\.]+)?\)$/);
    if (!parts) return undefined;
    const toHex = (c: string) => ('0' + parseInt(c).toString(16)).slice(-2);
    return `${toHex(parts[1])}${toHex(parts[2])}${toHex(parts[3])}`;
  }

  private base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binary_string = window.atob(base64);
    const len = binary_string.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
  }

  private getImageDimensionsFromBuffer(buffer: ArrayBuffer): { width: number, height: number } | null {
    const view = new DataView(buffer);
    try {
      if (view.getUint16(0) === 0xFFD8) { // JPEG marker
        let offset = 2;
        while (offset < buffer.byteLength) {
          if (view.getUint8(offset) !== 0xFF) {
            offset++;
            continue;
          }
          const marker = view.getUint8(offset + 1);
          if (marker === 0xC0 || marker === 0xC2) { // SOF0, SOF2
            const height = view.getUint16(offset + 5);
            const width = view.getUint16(offset + 7);
            return { width, height };
          }
          offset += 2 + view.getUint16(offset + 2);
        }
      } else if (view.getUint32(0) === 0x89504E47) { // PNG marker
        const width = view.getUint32(16);
        const height = view.getUint32(20);
        return { width, height };
      }
    } catch (e) {
      console.error("Failed to parse image dimensions from buffer:", e);
    }
    return null;
  }

  private detectMimeFromHeader(u8: Uint8Array): string | null {
    if (!u8 || u8.length < 8) return null;
    if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4E && u8[3] === 0x47) return 'image/png';
    if (u8[0] === 0xFF && u8[1] === 0xD8 && u8[2] === 0xFF) return 'image/jpeg';
    if (u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46) return 'image/gif';
    if (u8[0] === 0x42 && u8[1] === 0x4D) return 'image/bmp';
    if (u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 &&
      u8[8] === 0x57 && u8[9] === 0x45 && u8[10] === 0x42 && u8[11] === 0x50) return 'image/webp';
    const textStart = new TextDecoder().decode(u8.slice(0, 64)).trim();
    if (textStart.indexOf('<svg') !== -1) return 'image/svg+xml';
    return null;
  }

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

  private escapeRegExp(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  // --- 解析器 ---
  private parseSyntaxHighlightedCode(codeElement: HTMLElement, defaultFont: any, defaultSize: number): TextRun[] {
    const runs: TextRun[] = [];
    const nodes = Array.from(codeElement.childNodes);
    let hasContent = false;
    const codeFont = { name: 'Courier New' };
    nodes.forEach(node => {
      if (node.nodeType === Node.TEXT_NODE && node.textContent) {
        const style = window.getComputedStyle(codeElement);
        const segments = node.textContent.split('\n');
        segments.forEach((segment, segmentIndex) => {
          if (segment) {
            runs.push(new TextRun({ text: segment, font: codeFont, color: this.rgbToHex(style.color), size: defaultSize }));
            hasContent = true;
          }
          if (segmentIndex < segments.length - 1) {
            runs.push(new TextRun({ break: 1 }));
            hasContent = true;
          }
        });
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        const style = window.getComputedStyle(el);
        const color = this.rgbToHex(style.color);
        const segments = el.textContent?.split('\n') || [''];
        segments.forEach((segment, segmentIndex) => {
          if (segment) {
            runs.push(new TextRun({ text: segment, font: codeFont, color: color, size: defaultSize }));
            hasContent = true;
          }
          if (segmentIndex < segments.length - 1) {
            runs.push(new TextRun({ break: 1 }));
            hasContent = true;
          }
        });
      }
    });
    if (!hasContent && codeElement.textContent?.includes('\n')) {
      runs.push(new TextRun({ text: '', break: 1 }));
    }
    return runs;
  }

  private async parseListElement(listEl: HTMLUListElement | HTMLOListElement, level: number, bodyBgColor: string, sourcePath: string): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    const listItems = Array.from(listEl.children).filter(child => child.tagName === 'LI');
    const listType = listEl.tagName === 'OL' ? 'number' : 'bullet';
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
        const style = window.getComputedStyle(liElement);
        const paragraphOptions: any = {};
        const bgColor = this.rgbToHex(style.backgroundColor);
        if (bgColor && bgColor !== bodyBgColor) {
          paragraphOptions.shading = { type: ShadingType.CLEAR, fill: bgColor, color: "auto" };
        }
        const children = await this.parseInlineElements(contentContainer, sourcePath);
        if (children.length > 0) {
          const paragraphProperties: any = { ...paragraphOptions, children, spacing: { after: 100 } };
          if (listType === 'bullet') {
            paragraphProperties.bullet = { level };
          } else {
            paragraphProperties.numbering = { reference: "default-numbering", level };
          }
          paragraphs.push(new Paragraph(paragraphProperties));
        }
      }
      if (nestedList) {
        paragraphs.push(...await this.parseListElement(nestedList, level + 1, bodyBgColor, sourcePath));
      }
    }
    return paragraphs;
  }

  private async parseTableElement(tableEl: HTMLTableElement, bodyBgColor: string, sourcePath: string): Promise<Table> {
    const rows = Array.from(tableEl.querySelectorAll('tr'));
    const docxRows: TableRow[] = [];
    for (const row of rows) {
      const cells = Array.from(row.querySelectorAll('th, td'));
      const docxCells: TableCell[] = [];
      for (const cell of cells) {
        const cellContent = await this.htmlToDocxObjects(cell as HTMLElement, bodyBgColor, false, 0, sourcePath);
        docxCells.push(new TableCell({ children: cellContent }));
      }
      docxRows.push(new TableRow({ children: docxCells }));
    }
    return new Table({ rows: docxRows, width: { size: 100, type: WidthType.PERCENTAGE } });
  }

  private async parseInlineElements(element: HTMLElement, sourcePath: string): Promise<(TextRun | ExternalHyperlink | ImageRun)[]> {
    const runs: (TextRun | ExternalHyperlink | ImageRun)[] = [];
    const children = Array.from(element.childNodes);
    const mainFont = { name: 'Times New Roman' };
    const codeFont = { name: 'Courier New' };

    for (const child of children) {
      if (child.nodeType === Node.TEXT_NODE) {
        if (!child.textContent) continue;
        const style = window.getComputedStyle(element);
        const textContent = child.textContent.trim();
        if (textContent.length === 0) continue;
        const runOptions: any = { text: textContent };
        runOptions.color = this.rgbToHex(style.color);
        runOptions.size = this.pxToHalfPoints(style.fontSize);
        runOptions.font = mainFont;
        if (parseInt(style.fontWeight) >= 600) runOptions.bold = true;
        if (style.fontStyle === 'italic') runOptions.italics = true;
        runs.push(new TextRun(runOptions));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        const style = window.getComputedStyle(el);
        const runOptions: any = { text: el.textContent?.trim() || '' };
        runOptions.color = this.rgbToHex(style.color);
        runOptions.size = this.pxToHalfPoints(style.fontSize);
        runOptions.font = mainFont;
        switch (el.tagName.toUpperCase()) {
          case 'A':
            runs.push(new ExternalHyperlink({ link: el.getAttribute('href') || '', children: [new TextRun({ text: el.textContent?.trim() || '', style: "Hyperlink", color: this.rgbToHex(style.color) || '0563C1', underline: {} })], }));
            break;
          case 'DEL': runOptions.strike = true; runs.push(new TextRun(runOptions)); break;
          case 'STRONG':
          case 'B': runOptions.bold = true; runs.push(new TextRun(runOptions)); break;
          case 'EM':
          case 'I': runOptions.italics = true; runs.push(new TextRun(runOptions)); break;
          case 'U': runOptions.underline = {}; runs.push(new TextRun(runOptions)); break;
          case 'CODE':
            runOptions.style = "SourceCode";
            runOptions.shading = { type: ShadingType.CLEAR, fill: 'D3D3D3', color: "auto" };
            runOptions.font = codeFont;
            runs.push(new TextRun(runOptions));
            break;
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
          console.error(this.i18n.t("LOCAL_IMAGE_NOT_FOUND", pathFromEmbed), file);
          return null;
        }
        
        try {
          buffer = await this.app.vault.readBinary(file);
          imageExtension = file.extension;
        } catch (readError) {
          new Notice(this.i18n.t("FILE_READ_FAILED", file.path));
          console.error(this.i18n.t("FILE_READ_FAILED", file.path), readError);
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
      console.error(this.i18n.t("IMAGE_PROCESSING_ERROR", pathForNotice), error);
      return null;
    }
  }

  private async htmlToDocxObjects(element: HTMLElement, bodyBgColor: string, isTopLevel: boolean, indentLevel: number, sourcePath: string): Promise<(Paragraph | Table)[]> {
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

    if (!children.some(c => c.nodeType === Node.ELEMENT_NODE) && element.textContent?.trim()) {
      const inlineChildren = await this.parseInlineElements(element, sourcePath);
      if (inlineChildren.length > 0) { return [new Paragraph({ children: inlineChildren, font: mainFont })]; }
      return [];
    }

    for (let i = 0; i < children.length; i++) {
        const child = children[i];
        if (child.nodeType !== Node.ELEMENT_NODE) continue;

        const el = child as HTMLElement;
        const tagName = el.tagName.toUpperCase();
        let currentParagraphOptions: any = { ...paragraphStyles, font: mainFont };

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
                docxObjects.push(...await this.parseListElement(el as HTMLUListElement | HTMLOListElement, 0, bodyBgColor, sourcePath));
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
        const tagName = el.tagName.toUpperCase();

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
      shading: { type: ShadingType.CLEAR, fill: 'F0F0F0', color: "auto" },
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
  
  private async parseListElementForQuote(listEl: HTMLUListElement | HTMLOListElement, indentLevel: number, sourcePath: string): Promise<Paragraph[]> {
      const paragraphs: Paragraph[] = [];
      const listItems = Array.from(listEl.children).filter(child => child.tagName === 'LI');
      const listType = listEl.tagName === 'OL' ? 'number' : 'bullet';
      const mainFont = { name: 'Times New Roman' };
      
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
                  indent: { left: 200 * (indentLevel + 1) },
                  font: mainFont
              };
              
              if (listType === 'bullet') {
                  paragraphProperties.bullet = { level: indentLevel };
              } else {
                  paragraphProperties.numbering = { reference: "default-numbering", level: indentLevel };
              }
              
              paragraphs.push(new Paragraph(paragraphProperties));
          }

          if (nestedList) {
              paragraphs.push(...await this.parseListElementForQuote(nestedList, indentLevel + 1, sourcePath));
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
          shading: { type: ShadingType.CLEAR, fill: 'D3D3D3', color: "auto" },
          font: codeFont
      };
      
      return new Paragraph({ ...paragraphOptions, children: runs });
  }

  // --- 文件保存与主逻辑 ---

  private async saveFile(filePath: string, data: ArrayBuffer) {
    try {
      await this.app.vault.adapter.writeBinary(filePath, data);
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
    Object.assign(tempDiv.style, {
      position: 'absolute',
      top: '-9999px',
      left: '-9999px',
      width: '800px',
      height: 'auto'
    });

    try {
      document.body.appendChild(tempDiv);
      const markdownContent = await this.app.vault.read(activeFile);
      const sourcePath = activeFile.path;

      await MarkdownRenderer.render(this.app, markdownContent, tempDiv, sourcePath, this);

      const bodyBgColor = this.rgbToHex(window.getComputedStyle(document.body).backgroundColor);

      const docxObjects = await this.htmlToDocxObjects(tempDiv, bodyBgColor, true, 0, sourcePath);

      const title = activeFile.basename;
      const titleParagraph = new Paragraph({ text: title, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { after: 400 }, font: { name: 'Times New Roman' } });

      const doc = new Document({
        numbering: {
          config: [{
            reference: "default-numbering",
            levels: [
              { level: 0, format: "decimal", text: "%1.", alignment: AlignmentType.START, indent: { left: 720, hanging: 360 } },
              { level: 1, format: "decimal", text: "%1.%2.", alignment: AlignmentType.START, indent: { left: 1440, hanging: 360 } },
              { level: 2, format: "decimal", text: "%1.%2.%3.", alignment: AlignmentType.START, indent: { left: 2160, hanging: 360 } },
              { level: 3, format: "decimal", text: "%1.%2.%3.%4.", alignment: AlignmentType.START, indent: { left: 2880, hanging: 360 } },
              { level: 4, format: "decimal", text: "%1.%2.%3.%4.%5.", alignment: AlignmentType.START, indent: { left: 3600, hanging: 360 } },
            ],
          }, ],
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
      console.error(this.i18n.t("EXPORT_FAILED"), error);
    } finally {
      if (document.body.contains(tempDiv)) {
        document.body.removeChild(tempDiv);
      }
    }
  }
}
