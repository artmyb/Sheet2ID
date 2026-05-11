const path = require("path");
const fs = require("fs/promises");
const fontkit = require("@pdf-lib/fontkit");
const {
  PDFDocument,
  PageSizes,
  StandardFonts,
  clip,
  closePath,
  endPath,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
} = require("pdf-lib");
const { clamp, getRowValue, hexToRgb, mmToPt, renderTemplate } = require("./utils");

const DEFAULT_FONT_KEY = "helvetica";
const DEFAULT_BOLD_FONT_KEY = "helvetica-bold";
const PROJECT_FONT_DIR = path.resolve(__dirname, "../fonts");

const UNICODE_SANS_REGULAR_CANDIDATES = [
  path.join(PROJECT_FONT_DIR, "NotoSans-Regular.ttf"),
  "C:\\Windows\\Fonts\\NotoSans-Regular.ttf",
  "C:\\Windows\\Fonts\\DejaVuSans.ttf",
  "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
];

const UNICODE_SANS_BOLD_CANDIDATES = [
  path.join(PROJECT_FONT_DIR, "NotoSans-Bold.ttf"),
  "C:\\Windows\\Fonts\\NotoSans-Bold.ttf",
  "C:\\Windows\\Fonts\\DejaVuSans-Bold.ttf",
  "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
];

const UNICODE_SANS_ITALIC_CANDIDATES = [
  path.join(PROJECT_FONT_DIR, "NotoSans-Italic.ttf"),
  "C:\\Windows\\Fonts\\NotoSans-Italic.ttf",
  "C:\\Windows\\Fonts\\DejaVuSans-Oblique.ttf",
  "/usr/share/fonts/truetype/noto/NotoSans-Italic.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
];

const UNICODE_SANS_BOLD_ITALIC_CANDIDATES = [
  path.join(PROJECT_FONT_DIR, "NotoSans-BoldItalic.ttf"),
  "C:\\Windows\\Fonts\\NotoSans-BoldItalic.ttf",
  "C:\\Windows\\Fonts\\DejaVuSans-BoldOblique.ttf",
  "/usr/share/fonts/truetype/noto/NotoSans-BoldItalic.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSans-BoldItalic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf",
];

const SANS_REGULAR_CANDIDATES = [
  ...UNICODE_SANS_REGULAR_CANDIDATES,
  "C:\\Windows\\Fonts\\arial.ttf",
];

const SANS_BOLD_CANDIDATES = [
  ...UNICODE_SANS_BOLD_CANDIDATES,
  "C:\\Windows\\Fonts\\arialbd.ttf",
];

const SANS_ITALIC_CANDIDATES = [
  ...UNICODE_SANS_ITALIC_CANDIDATES,
  "C:\\Windows\\Fonts\\ariali.ttf",
];

const SANS_BOLD_ITALIC_CANDIDATES = [
  ...UNICODE_SANS_BOLD_ITALIC_CANDIDATES,
  "C:\\Windows\\Fonts\\arialbi.ttf",
];

const SERIF_REGULAR_CANDIDATES = [
  "C:\\Windows\\Fonts\\times.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
  "/usr/share/fonts/truetype/noto/NotoSerif-Regular.ttf",
];

const SERIF_BOLD_CANDIDATES = [
  "C:\\Windows\\Fonts\\timesbd.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSerif-Bold.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
  "/usr/share/fonts/truetype/noto/NotoSerif-Bold.ttf",
];

const SERIF_ITALIC_CANDIDATES = [
  "C:\\Windows\\Fonts\\timesi.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Italic.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSerif-Italic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf",
  "/usr/share/fonts/truetype/noto/NotoSerif-Italic.ttf",
];

const SERIF_BOLD_ITALIC_CANDIDATES = [
  "C:\\Windows\\Fonts\\timesbi.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSerif-BoldItalic.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationSerif-BoldItalic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf",
  "/usr/share/fonts/truetype/noto/NotoSerif-BoldItalic.ttf",
];

const MONO_REGULAR_CANDIDATES = [
  "C:\\Windows\\Fonts\\cour.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationMono-Regular.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansMono-Regular.ttf",
];

const MONO_BOLD_CANDIDATES = [
  "C:\\Windows\\Fonts\\courbd.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationMono-Bold.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansMono-Bold.ttf",
];

const MONO_ITALIC_CANDIDATES = [
  "C:\\Windows\\Fonts\\couri.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Oblique.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationMono-Italic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationMono-Italic.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansMono-Italic.ttf",
];

const MONO_BOLD_ITALIC_CANDIDATES = [
  "C:\\Windows\\Fonts\\courbi.ttf",
  "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-BoldOblique.ttf",
  "/usr/share/fonts/truetype/liberation2/LiberationMono-BoldItalic.ttf",
  "/usr/share/fonts/truetype/liberation/LiberationMono-BoldItalic.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansMono-BoldItalic.ttf",
];

const TAHOMA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\tahoma.ttf", ...SANS_REGULAR_CANDIDATES];
const TAHOMA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\tahomabd.ttf", ...SANS_BOLD_CANDIDATES];
const TREBUCHET_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\trebuc.ttf", ...SANS_REGULAR_CANDIDATES];
const TREBUCHET_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\trebucbd.ttf", ...SANS_BOLD_CANDIDATES];
const TREBUCHET_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\trebucit.ttf", ...SANS_ITALIC_CANDIDATES];
const TREBUCHET_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\trebucbi.ttf", ...SANS_BOLD_ITALIC_CANDIDATES];
const GEORGIA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\georgia.ttf", ...SERIF_REGULAR_CANDIDATES];
const GEORGIA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\georgiab.ttf", ...SERIF_BOLD_CANDIDATES];
const GEORGIA_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\georgiai.ttf", ...SERIF_ITALIC_CANDIDATES];
const GEORGIA_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\georgiaz.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const PALATINO_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\pala.ttf", ...SERIF_REGULAR_CANDIDATES];
const PALATINO_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\palab.ttf", ...SERIF_BOLD_CANDIDATES];
const PALATINO_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\palai.ttf", ...SERIF_ITALIC_CANDIDATES];
const PALATINO_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\palabi.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const CAMBRIA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\cambria.ttf", "C:\\Windows\\Fonts\\cambria.ttc", ...SERIF_REGULAR_CANDIDATES];
const CAMBRIA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\cambriab.ttf", "C:\\Windows\\Fonts\\cambria.ttc", ...SERIF_BOLD_CANDIDATES];
const CAMBRIA_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\cambriai.ttf", "C:\\Windows\\Fonts\\cambria.ttc", ...SERIF_ITALIC_CANDIDATES];
const CAMBRIA_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\cambriaz.ttf", "C:\\Windows\\Fonts\\cambria.ttc", ...SERIF_BOLD_ITALIC_CANDIDATES];
const BOOK_ANTIQUA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\bkant.ttf", ...SERIF_REGULAR_CANDIDATES];
const BOOK_ANTIQUA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\bkantb.ttf", ...SERIF_BOLD_CANDIDATES];
const BOOK_ANTIQUA_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\bkanti.ttf", ...SERIF_ITALIC_CANDIDATES];
const BOOK_ANTIQUA_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\bkantbi.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const CONSTANTIA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\constan.ttf", ...SERIF_REGULAR_CANDIDATES];
const CONSTANTIA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\constanb.ttf", ...SERIF_BOLD_CANDIDATES];
const CONSTANTIA_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\constani.ttf", ...SERIF_ITALIC_CANDIDATES];
const CONSTANTIA_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\constanz.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const GARAMOND_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\gara.ttf", ...SERIF_REGULAR_CANDIDATES];
const GARAMOND_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\garabd.ttf", ...SERIF_BOLD_CANDIDATES];
const GARAMOND_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\garait.ttf", ...SERIF_ITALIC_CANDIDATES];
const GARAMOND_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\garaiti.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const ROCKWELL_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\rock.ttf", ...SERIF_REGULAR_CANDIDATES];
const ROCKWELL_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\rockb.ttf", ...SERIF_BOLD_CANDIDATES];
const ROCKWELL_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\rocki.ttf", ...SERIF_ITALIC_CANDIDATES];
const ROCKWELL_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\rockbi.ttf", ...SERIF_BOLD_ITALIC_CANDIDATES];
const ARIAL_NARROW_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\arialn.ttf", ...SANS_REGULAR_CANDIDATES];
const ARIAL_NARROW_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\arialnb.ttf", ...SANS_BOLD_CANDIDATES];
const ARIAL_NARROW_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\arialni.ttf", ...SANS_ITALIC_CANDIDATES];
const ARIAL_NARROW_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\arialnbi.ttf", ...SANS_BOLD_ITALIC_CANDIDATES];
const CANDARA_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\candara.ttf", ...SANS_REGULAR_CANDIDATES];
const CANDARA_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\candarab.ttf", ...SANS_BOLD_CANDIDATES];
const CANDARA_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\candarai.ttf", ...SANS_ITALIC_CANDIDATES];
const CANDARA_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\candaraz.ttf", ...SANS_BOLD_ITALIC_CANDIDATES];
const CORBEL_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\corbel.ttf", ...SANS_REGULAR_CANDIDATES];
const CORBEL_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\corbelb.ttf", ...SANS_BOLD_CANDIDATES];
const CORBEL_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\corbeli.ttf", ...SANS_ITALIC_CANDIDATES];
const CORBEL_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\corbelz.ttf", ...SANS_BOLD_ITALIC_CANDIDATES];
const CENTURY_GOTHIC_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\gothic.ttf", ...SANS_REGULAR_CANDIDATES];
const CENTURY_GOTHIC_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\gothicb.ttf", ...SANS_BOLD_CANDIDATES];
const CENTURY_GOTHIC_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\gothici.ttf", ...SANS_ITALIC_CANDIDATES];
const CENTURY_GOTHIC_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\gothicbi.ttf", ...SANS_BOLD_ITALIC_CANDIDATES];
const FRANKLIN_GOTHIC_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\framd.ttf", ...SANS_REGULAR_CANDIDATES];
const FRANKLIN_GOTHIC_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\framdit.ttf", ...SANS_ITALIC_CANDIDATES];
const COMIC_SANS_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\comic.ttf", ...SANS_REGULAR_CANDIDATES];
const COMIC_SANS_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\comicbd.ttf", ...SANS_BOLD_CANDIDATES];
const IMPACT_CANDIDATES = ["C:\\Windows\\Fonts\\impact.ttf", ...SANS_REGULAR_CANDIDATES];
const CONSOLAS_REGULAR_CANDIDATES = ["C:\\Windows\\Fonts\\consola.ttf", ...MONO_REGULAR_CANDIDATES];
const CONSOLAS_BOLD_CANDIDATES = ["C:\\Windows\\Fonts\\consolab.ttf", ...MONO_BOLD_CANDIDATES];
const CONSOLAS_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\consolai.ttf", ...MONO_ITALIC_CANDIDATES];
const CONSOLAS_BOLD_ITALIC_CANDIDATES = ["C:\\Windows\\Fonts\\consolaz.ttf", ...MONO_BOLD_ITALIC_CANDIDATES];
const LUCIDA_CONSOLE_CANDIDATES = ["C:\\Windows\\Fonts\\lucon.ttf", ...MONO_REGULAR_CANDIDATES];

const FONT_DEFINITIONS = {
  "helvetica": { candidates: SANS_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "helvetica-bold": { candidates: SANS_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "helvetica-oblique": { candidates: SANS_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "helvetica-boldoblique": { candidates: SANS_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "times-roman": { candidates: SERIF_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "times-bold": { candidates: SERIF_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "times-italic": { candidates: SERIF_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "times-bolditalic": { candidates: SERIF_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "courier": { candidates: MONO_REGULAR_CANDIDATES, fallback: StandardFonts.Courier },
  "courier-bold": { candidates: MONO_BOLD_CANDIDATES, fallback: StandardFonts.CourierBold },
  "courier-oblique": { candidates: MONO_ITALIC_CANDIDATES, fallback: StandardFonts.CourierOblique },
  "courier-boldoblique": { candidates: MONO_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.CourierBoldOblique },
  "arial": { candidates: SANS_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "arial bold": { candidates: SANS_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "arial italic": { candidates: SANS_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "arial bold italic": { candidates: SANS_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "segoe ui": { candidates: ["C:\\Windows\\Fonts\\segoeui.ttf", ...SANS_REGULAR_CANDIDATES], fallback: StandardFonts.Helvetica },
  "segoe ui bold": { candidates: ["C:\\Windows\\Fonts\\segoeuib.ttf", ...SANS_BOLD_CANDIDATES], fallback: StandardFonts.HelveticaBold },
  "segoe ui italic": { candidates: ["C:\\Windows\\Fonts\\segoeuii.ttf", ...SANS_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaOblique },
  "segoe ui bold italic": { candidates: ["C:\\Windows\\Fonts\\segoeuiz.ttf", ...SANS_BOLD_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaBoldOblique },
  "calibri": { candidates: ["C:\\Windows\\Fonts\\calibri.ttf", ...SANS_REGULAR_CANDIDATES], fallback: StandardFonts.Helvetica },
  "calibri bold": { candidates: ["C:\\Windows\\Fonts\\calibrib.ttf", ...SANS_BOLD_CANDIDATES], fallback: StandardFonts.HelveticaBold },
  "calibri italic": { candidates: ["C:\\Windows\\Fonts\\calibrii.ttf", ...SANS_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaOblique },
  "calibri bold italic": { candidates: ["C:\\Windows\\Fonts\\calibriz.ttf", ...SANS_BOLD_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaBoldOblique },
  "verdana": { candidates: ["C:\\Windows\\Fonts\\verdana.ttf", ...SANS_REGULAR_CANDIDATES], fallback: StandardFonts.Helvetica },
  "verdana bold": { candidates: ["C:\\Windows\\Fonts\\verdanab.ttf", ...SANS_BOLD_CANDIDATES], fallback: StandardFonts.HelveticaBold },
  "verdana italic": { candidates: ["C:\\Windows\\Fonts\\verdanai.ttf", ...SANS_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaOblique },
  "verdana bold italic": { candidates: ["C:\\Windows\\Fonts\\verdanaz.ttf", ...SANS_BOLD_ITALIC_CANDIDATES], fallback: StandardFonts.HelveticaBoldOblique },
  "tahoma": { candidates: TAHOMA_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "tahoma bold": { candidates: TAHOMA_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "tahoma italic": { candidates: SANS_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "tahoma bold italic": { candidates: SANS_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "trebuchet ms": { candidates: TREBUCHET_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "trebuchet ms bold": { candidates: TREBUCHET_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "trebuchet ms italic": { candidates: TREBUCHET_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "trebuchet ms bold italic": { candidates: TREBUCHET_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "georgia": { candidates: GEORGIA_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "georgia bold": { candidates: GEORGIA_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "georgia italic": { candidates: GEORGIA_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "georgia bold italic": { candidates: GEORGIA_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "palatino linotype": { candidates: PALATINO_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "palatino linotype bold": { candidates: PALATINO_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "palatino linotype italic": { candidates: PALATINO_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "palatino linotype bold italic": { candidates: PALATINO_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "cambria": { candidates: CAMBRIA_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "cambria bold": { candidates: CAMBRIA_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "cambria italic": { candidates: CAMBRIA_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "cambria bold italic": { candidates: CAMBRIA_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "book antiqua": { candidates: BOOK_ANTIQUA_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "book antiqua bold": { candidates: BOOK_ANTIQUA_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "book antiqua italic": { candidates: BOOK_ANTIQUA_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "book antiqua bold italic": { candidates: BOOK_ANTIQUA_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "constantia": { candidates: CONSTANTIA_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "constantia bold": { candidates: CONSTANTIA_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "constantia italic": { candidates: CONSTANTIA_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "constantia bold italic": { candidates: CONSTANTIA_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "garamond": { candidates: GARAMOND_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "garamond bold": { candidates: GARAMOND_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "garamond italic": { candidates: GARAMOND_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "garamond bold italic": { candidates: GARAMOND_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "rockwell": { candidates: ROCKWELL_REGULAR_CANDIDATES, fallback: StandardFonts.TimesRoman },
  "rockwell bold": { candidates: ROCKWELL_BOLD_CANDIDATES, fallback: StandardFonts.TimesRomanBold },
  "rockwell italic": { candidates: ROCKWELL_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanItalic },
  "rockwell bold italic": { candidates: ROCKWELL_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.TimesRomanBoldItalic },
  "arial narrow": { candidates: ARIAL_NARROW_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "arial narrow bold": { candidates: ARIAL_NARROW_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "arial narrow italic": { candidates: ARIAL_NARROW_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "arial narrow bold italic": { candidates: ARIAL_NARROW_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "candara": { candidates: CANDARA_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "candara bold": { candidates: CANDARA_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "candara italic": { candidates: CANDARA_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "candara bold italic": { candidates: CANDARA_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "corbel": { candidates: CORBEL_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "corbel bold": { candidates: CORBEL_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "corbel italic": { candidates: CORBEL_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "corbel bold italic": { candidates: CORBEL_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "century gothic": { candidates: CENTURY_GOTHIC_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "century gothic bold": { candidates: CENTURY_GOTHIC_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "century gothic italic": { candidates: CENTURY_GOTHIC_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "century gothic bold italic": { candidates: CENTURY_GOTHIC_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaBoldOblique },
  "franklin gothic medium": { candidates: FRANKLIN_GOTHIC_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "franklin gothic medium italic": { candidates: FRANKLIN_GOTHIC_ITALIC_CANDIDATES, fallback: StandardFonts.HelveticaOblique },
  "comic sans ms": { candidates: COMIC_SANS_REGULAR_CANDIDATES, fallback: StandardFonts.Helvetica },
  "comic sans ms bold": { candidates: COMIC_SANS_BOLD_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "impact": { candidates: IMPACT_CANDIDATES, fallback: StandardFonts.HelveticaBold },
  "consolas": { candidates: CONSOLAS_REGULAR_CANDIDATES, fallback: StandardFonts.Courier },
  "consolas bold": { candidates: CONSOLAS_BOLD_CANDIDATES, fallback: StandardFonts.CourierBold },
  "consolas italic": { candidates: CONSOLAS_ITALIC_CANDIDATES, fallback: StandardFonts.CourierOblique },
  "consolas bold italic": { candidates: CONSOLAS_BOLD_ITALIC_CANDIDATES, fallback: StandardFonts.CourierBoldOblique },
  "lucida console": { candidates: LUCIDA_CONSOLE_CANDIDATES, fallback: StandardFonts.Courier },
};

const FONT_ALIASES = {
  "times": "times-roman",
  "times roman": "times-roman",
  "times new roman": "times-roman",
  "times bold": "times-bold",
  "times new roman bold": "times-bold",
  "times italic": "times-italic",
  "times new roman italic": "times-italic",
  "times bold italic": "times-bolditalic",
  "times new roman bold italic": "times-bolditalic",
  "serif": "times-roman",
  "sans serif": "helvetica",
  "sans-serif": "helvetica",
  "courier new": "courier",
  "courier new bold": "courier-bold",
  "courier new italic": "courier-oblique",
  "courier new bold italic": "courier-boldoblique",
  "dejavu sans": "helvetica",
  "dejavu sans bold": "helvetica-bold",
  "dejavu sans italic": "helvetica-oblique",
  "dejavu sans bold italic": "helvetica-boldoblique",
  "liberation sans": "helvetica",
  "liberation sans bold": "helvetica-bold",
  "liberation sans italic": "helvetica-oblique",
  "liberation sans bold italic": "helvetica-boldoblique",
  "noto sans": "helvetica",
  "noto sans bold": "helvetica-bold",
  "noto sans italic": "helvetica-oblique",
  "noto sans bold italic": "helvetica-boldoblique",
  "dejavu serif": "times-roman",
  "dejavu serif bold": "times-bold",
  "dejavu serif italic": "times-italic",
  "dejavu serif bold italic": "times-bolditalic",
  "liberation serif": "times-roman",
  "liberation serif bold": "times-bold",
  "liberation serif italic": "times-italic",
  "liberation serif bold italic": "times-bolditalic",
  "noto serif": "times-roman",
  "noto serif bold": "times-bold",
  "noto serif italic": "times-italic",
  "noto serif bold italic": "times-bolditalic",
  "dejavu sans mono": "courier",
  "dejavu sans mono bold": "courier-bold",
  "dejavu sans mono italic": "courier-oblique",
  "dejavu sans mono bold italic": "courier-boldoblique",
  "liberation mono": "courier",
  "liberation mono bold": "courier-bold",
  "liberation mono italic": "courier-oblique",
  "liberation mono bold italic": "courier-boldoblique",
  "noto sans mono": "courier",
  "noto sans mono bold": "courier-bold",
  "noto sans mono italic": "courier-oblique",
  "noto sans mono bold italic": "courier-boldoblique",
  "monospace": "courier",
};

async function generatePdf(cards, config) {
  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  const fontCache = await embedFonts(pdfDoc, config);
  const imageCache = new Map();

  const pageSize = resolvePageSize(config);
  const grid = resolveGrid(config, pageSize);
  const cardsPerPage = grid.columns * grid.rows;

  if (cardsPerPage < 1) {
    throw new Error("The current page and card size settings leave no room for cards.");
  }

  for (let index = 0; index < cards.length; index += 1) {
    if (index % cardsPerPage === 0) {
      pdfDoc.addPage([pageSize.width, pageSize.height]);
    }

    const page = pdfDoc.getPages()[Math.floor(index / cardsPerPage)];
    const slot = index % cardsPerPage;
    const cardRect = resolveCardRect(slot, grid, pageSize, config);
    await drawCard(page, pdfDoc, cards[index], cardRect, config, fontCache, imageCache);
  }

  return pdfDoc.save();
}

async function drawCard(page, pdfDoc, card, cardRect, config, fontCache, imageCache) {
  drawCardBase(page, cardRect, config);

  if (card.backgroundAsset) {
    const background = await getEmbeddedImage(pdfDoc, card.backgroundAsset, imageCache);
    drawImageInBox(page, background, cardRect, "cover");
  }

  if (config.card.borderWidthMm > 0) {
    page.drawRectangle({
      x: cardRect.x,
      y: cardRect.y,
      width: cardRect.width,
      height: cardRect.height,
      borderWidth: mmToPt(config.card.borderWidthMm),
      borderColor: hexToRgb(config.card.borderColor, "#D6D0C4"),
    });
  }

  for (const element of config.elements) {
    if (element.type === "photo") {
      await drawPhotoElement(page, pdfDoc, card, cardRect, element, config, fontCache, imageCache);
      continue;
    }

    if (element.type === "text") {
      drawTextElement(page, card, cardRect, element, fontCache);
    }
  }
}

function drawCardBase(page, cardRect, config) {
  page.drawRectangle({
    x: cardRect.x,
    y: cardRect.y,
    width: cardRect.width,
    height: cardRect.height,
    color: hexToRgb(config.card.fillColor, "#FFFFFF"),
  });
}

async function drawPhotoElement(page, pdfDoc, card, cardRect, element, config, fontCache, imageCache) {
  const box = resolveElementBox(cardRect, element);

  if (element.fillColor) {
    page.drawRectangle({
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      color: hexToRgb(element.fillColor, "#F3F0EA"),
    });
  }

  if (card.photoAsset) {
    const image = await getEmbeddedImage(pdfDoc, card.photoAsset, imageCache);
    drawImageInBox(page, image, box, element.fit);
  } else {
    const label = config.photos.placeholderLabel || "NO PHOTO";
    const prepared = prepareTextForRendering(label, [
      fontCache[DEFAULT_BOLD_FONT_KEY] || fontCache[DEFAULT_FONT_KEY],
      fontCache[DEFAULT_FONT_KEY],
    ]);
    const size = clamp(box.height / 6.5, 7, 12);
    const textWidth = prepared.font.widthOfTextAtSize(prepared.text, size);
    page.drawText(prepared.text, {
      x: box.x + Math.max(0, (box.width - textWidth) / 2),
      y: box.y + Math.max(0, (box.height - size) / 2),
      size,
      font: prepared.font,
      color: hexToRgb("#8B8577", "#8B8577"),
    });
  }

  if (element.borderWidthMm > 0) {
    page.drawRectangle({
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      borderWidth: mmToPt(element.borderWidthMm),
      borderColor: hexToRgb(element.borderColor || "#FFFFFF", "#FFFFFF"),
    });
  }
}

function drawTextElement(page, card, cardRect, element, fontCache) {
  let text = getElementText(card.row, element);
  if (!text && element.hideWhenEmpty) {
    return;
  }

  if (element.uppercase) {
    text = text.toUpperCase();
  }

  const fontKey = resolveFontKey(element.font);
  const prepared = prepareTextForRendering(text, resolvePreferredFonts(fontCache, fontKey));
  const font = prepared.font;
  text = prepared.text;
  const box = resolveElementBox(cardRect, element);
  const padding = mmToPt(element.paddingMm);

  const fitted = fitTextToBox(text, font, {
    width: Math.max(0, box.width - padding * 2),
    height: Math.max(0, box.height - padding * 2),
    fontSize: element.fontSize,
    minFontSize: element.minFontSize,
    lineHeight: element.lineHeight,
    maxLines: element.maxLines,
  });

  if (!fitted.lines.length) {
    return;
  }

  const contentHeight = fitted.lines.length * fitted.lineHeight;
  let startY = box.y + box.height - padding - fitted.size;

  if (element.valign === "center") {
    startY = box.y + (box.height + contentHeight) / 2 - fitted.size;
  } else if (element.valign === "bottom") {
    startY = box.y + padding + (fitted.lines.length - 1) * fitted.lineHeight;
  }

  fitted.lines.forEach((line, index) => {
    const lineWidth = font.widthOfTextAtSize(line, fitted.size);
    let x = box.x + padding;

    if (element.align === "center") {
      x = box.x + (box.width - lineWidth) / 2;
    } else if (element.align === "right") {
      x = box.x + box.width - padding - lineWidth;
    }

    page.drawText(line, {
      x,
      y: startY - index * fitted.lineHeight,
      size: fitted.size,
      font,
      color: hexToRgb(element.color, "#111111"),
    });
  });
}

function resolvePreferredFonts(fontCache, fontKey) {
  const candidates = [fontCache[fontKey]];

  if (String(fontKey || "").includes("bold")) {
    candidates.push(fontCache[DEFAULT_BOLD_FONT_KEY]);
  }

  candidates.push(fontCache[DEFAULT_FONT_KEY]);

  return candidates.filter((font, index, list) => font && list.indexOf(font) === index);
}

function prepareTextForRendering(text, candidateFonts) {
  const normalizedText = String(text || "");
  const fonts = candidateFonts.filter(Boolean);
  const fallbackFont = fonts[0];

  if (!fallbackFont) {
    throw new Error("No PDF font was available for rendering text.");
  }

  for (const font of fonts) {
    if (canFontEncode(font, normalizedText)) {
      return { font, text: normalizedText };
    }
  }

  const fallbackText = sanitizeUnsupportedText(normalizedText);

  for (const font of fonts) {
    if (canFontEncode(font, fallbackText)) {
      return { font, text: fallbackText };
    }
  }

  return {
    font: fallbackFont,
    text: fallbackText.replace(/[^\x20-\x7E\r\n\t]/g, "?"),
  };
}

function canFontEncode(font, text) {
  try {
    font.encodeText(String(text || ""));
    return true;
  } catch {
    return false;
  }
}

function sanitizeUnsupportedText(text) {
  const turkishAsciiMap = {
    "Ç": "C",
    "ç": "c",
    "Ğ": "G",
    "ğ": "g",
    "İ": "I",
    "ı": "i",
    "Ö": "O",
    "ö": "o",
    "Ş": "S",
    "ş": "s",
    "Ü": "U",
    "ü": "u",
  };

  return String(text || "")
    .replace(/[ÇçĞğİıÖöŞşÜü]/g, (character) => turkishAsciiMap[character] || character)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\x20-\x7E\r\n\t]/g, "?");
}

function fitTextToBox(text, font, options) {
  const normalizedText = String(text || "").trim();
  if (!normalizedText) {
    return {
      lines: [],
      size: options.fontSize,
      lineHeight: options.fontSize * options.lineHeight,
    };
  }

  for (let size = options.fontSize; size >= options.minFontSize; size -= 0.5) {
    const lineHeight = size * options.lineHeight;
    const maxLinesByHeight = Math.max(1, Math.floor(options.height / lineHeight));
    const maxLines = Math.max(1, Math.min(options.maxLines, maxLinesByHeight));
    const lines = wrapText(normalizedText, font, size, options.width);

    if (lines.length <= maxLines) {
      return { lines, size, lineHeight };
    }
  }

  const size = options.minFontSize;
  const lineHeight = size * options.lineHeight;
  const maxLinesByHeight = Math.max(1, Math.floor(options.height / lineHeight));
  const maxLines = Math.max(1, Math.min(options.maxLines, maxLinesByHeight));
  const wrapped = wrapText(normalizedText, font, size, options.width);
  const lines = truncateLines(wrapped, maxLines, font, size, options.width);
  return { lines, size, lineHeight };
}

function wrapText(text, font, size, width) {
  const paragraphs = String(text).split(/\r?\n/);
  const lines = [];

  paragraphs.forEach((paragraph) => {
    const words = paragraph.split(/\s+/).filter(Boolean);
    if (!words.length) {
      lines.push("");
      return;
    }

    let current = "";

    for (const word of words) {
      const candidate = current ? `${current} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, size) <= width) {
        current = candidate;
        continue;
      }

      if (current) {
        lines.push(current);
      }

      if (font.widthOfTextAtSize(word, size) <= width) {
        current = word;
        continue;
      }

      const pieces = breakWord(word, font, size, width);
      lines.push(...pieces.slice(0, -1));
      current = pieces[pieces.length - 1];
    }

    if (current) {
      lines.push(current);
    }
  });

  return lines;
}

function breakWord(word, font, size, width) {
  const parts = [];
  let current = "";

  for (const character of word) {
    const candidate = current + character;
    if (!current || font.widthOfTextAtSize(candidate, size) <= width) {
      current = candidate;
      continue;
    }

    parts.push(current);
    current = character;
  }

  if (current) {
    parts.push(current);
  }

  return parts;
}

function truncateLines(lines, maxLines, font, size, width) {
  if (lines.length <= maxLines) {
    return lines;
  }

  const truncated = lines.slice(0, maxLines);
  let lastLine = truncated[maxLines - 1];

  while (lastLine && font.widthOfTextAtSize(`${lastLine}...`, size) > width) {
    lastLine = lastLine.slice(0, -1);
  }

  truncated[maxLines - 1] = lastLine ? `${lastLine}...` : "...";
  return truncated;
}

function getElementText(row, element) {
  if (element.template) {
    return renderTemplate(element.template, row);
  }

  if (element.field) {
    return `${element.prefix || ""}${getRowValue(row, element.field)}${element.suffix || ""}`.trim();
  }

  return `${element.prefix || ""}${element.value || ""}${element.suffix || ""}`.trim();
}

function resolveElementBox(cardRect, element) {
  const width = mmToPt(element.widthMm);
  const height = mmToPt(element.heightMm);
  return {
    x: cardRect.x + mmToPt(element.xMm),
    y: cardRect.y + cardRect.height - mmToPt(element.yMm) - height,
    width,
    height,
  };
}

function resolvePageSize(config) {
  const presetEntry = Object.entries(PageSizes).find(([name]) => name.toLowerCase() === config.page.size.toLowerCase());
  let width;
  let height;

  if (presetEntry) {
    [width, height] = presetEntry[1];
  } else {
    width = mmToPt(config.page.widthMm);
    height = mmToPt(config.page.heightMm);
  }

  if (config.page.orientation === "landscape") {
    return {
      width: Math.max(width, height),
      height: Math.min(width, height),
    };
  }

  return {
    width: Math.min(width, height),
    height: Math.max(width, height),
  };
}

function resolveGrid(config, pageSize) {
  const margin = mmToPt(config.page.marginMm);
  const gap = mmToPt(config.page.gapMm);
  const cardWidth = mmToPt(config.card.widthMm);
  const cardHeight = mmToPt(config.card.heightMm);
  const usableWidth = pageSize.width - margin * 2;
  const usableHeight = pageSize.height - margin * 2;
  const columns = config.layout.columns;
  const rows = config.layout.rows;
  const gridWidth = columns * cardWidth + Math.max(0, columns - 1) * gap;
  const gridHeight = rows * cardHeight + Math.max(0, rows - 1) * gap;

  return {
    margin,
    gap,
    cardWidth,
    cardHeight,
    columns,
    rows,
    offsetX: margin + (config.layout.centerOnPage ? Math.max(0, (usableWidth - gridWidth) / 2) : 0),
    offsetY: margin + (config.layout.centerOnPage ? Math.max(0, (usableHeight - gridHeight) / 2) : 0),
  };
}

function resolveCardRect(slot, grid, pageSize) {
  const column = slot % grid.columns;
  const row = Math.floor(slot / grid.columns);
  return {
    x: grid.offsetX + column * (grid.cardWidth + grid.gap),
    y: pageSize.height - grid.offsetY - grid.cardHeight - row * (grid.cardHeight + grid.gap),
    width: grid.cardWidth,
    height: grid.cardHeight,
  };
}

async function embedFonts(pdfDoc, config) {
  const requestedFonts = new Set([DEFAULT_FONT_KEY, DEFAULT_BOLD_FONT_KEY]);
  config.elements.forEach((element) => {
    if (element.type === "text") {
      requestedFonts.add(resolveFontKey(element.font));
    }
  });

  const fontEntries = await Promise.all(
    [...requestedFonts].map(async (fontKey) => [fontKey, await embedFont(pdfDoc, fontKey)])
  );

  return Object.fromEntries(fontEntries);
}

async function embedFont(pdfDoc, fontKey) {
  const normalizedKey = resolveFontKey(fontKey);
  const definition = FONT_DEFINITIONS[normalizedKey] || FONT_DEFINITIONS[DEFAULT_FONT_KEY];

  const embeddedRequestedFont = await embedFirstReadableFont(pdfDoc, definition.candidates || []);
  if (embeddedRequestedFont) {
    return embeddedRequestedFont;
  }

  const unicodeFallbackFont = await embedFirstReadableFont(pdfDoc, getUnicodeFallbackCandidates(normalizedKey));
  if (unicodeFallbackFont) {
    return unicodeFallbackFont;
  }

  return pdfDoc.embedFont(definition.fallback || StandardFonts.Helvetica);
}

async function embedFirstReadableFont(pdfDoc, candidates) {
  for (const candidatePath of candidates) {
    try {
      const fontBytes = await fs.readFile(candidatePath);
      return await pdfDoc.embedFont(fontBytes, { subset: true });
    } catch {
      continue;
    }
  }

  return null;
}

function getUnicodeFallbackCandidates(fontKey) {
  const normalized = String(fontKey || "").toLowerCase();
  const isBold = normalized.includes("bold");
  const isItalic = normalized.includes("italic") || normalized.includes("oblique");

  if (isBold && isItalic) {
    return UNICODE_SANS_BOLD_ITALIC_CANDIDATES;
  }

  if (isBold) {
    return UNICODE_SANS_BOLD_CANDIDATES;
  }

  if (isItalic) {
    return UNICODE_SANS_ITALIC_CANDIDATES;
  }

  return UNICODE_SANS_REGULAR_CANDIDATES;
}

function resolveFontKey(fontName) {
  const normalized = String(fontName || DEFAULT_FONT_KEY)
    .toLowerCase()
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (FONT_DEFINITIONS[normalized]) {
    return normalized;
  }

  return FONT_ALIASES[normalized] || DEFAULT_FONT_KEY;
}

async function getEmbeddedImage(pdfDoc, asset, imageCache) {
  if (imageCache.has(asset.id)) {
    return imageCache.get(asset.id);
  }

  const embeddedPromise = asset.extension === ".png" ? pdfDoc.embedPng(asset.buffer) : pdfDoc.embedJpg(asset.buffer);
  imageCache.set(asset.id, embeddedPromise);
  return embeddedPromise;
}

function drawImageInBox(page, image, box, fit = "cover") {
  if (fit === "contain") {
    const scaled = image.scaleToFit(box.width, box.height);
    page.drawImage(image, {
      x: box.x + (box.width - scaled.width) / 2,
      y: box.y + (box.height - scaled.height) / 2,
      width: scaled.width,
      height: scaled.height,
    });
    return;
  }

  const scale = Math.max(box.width / image.width, box.height / image.height);
  const width = image.width * scale;
  const height = image.height * scale;
  const x = box.x + (box.width - width) / 2;
  const y = box.y + (box.height - height) / 2;

  page.pushOperators(
    pushGraphicsState(),
    moveTo(box.x, box.y),
    lineTo(box.x, box.y + box.height),
    lineTo(box.x + box.width, box.y + box.height),
    lineTo(box.x + box.width, box.y),
    closePath(),
    clip(),
    endPath()
  );
  page.drawImage(image, { x, y, width, height });
  page.pushOperators(popGraphicsState());
}

module.exports = {
  generatePdf,
};
