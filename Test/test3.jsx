/**
 * InDesign → Photoshop / Illustrator link opener
 * Professional TypeScript implementation (ES2022)
 * -----------------------------------------------
 * 1.  Compile:  tsc openLink.ts --target ES2022 --module CommonJS
 * 2.  Drop generated  openLink.jsx  into InDesign Scripts Panel
 * 3.  Select any graphic frame and run – the linked file opens in the
 *     correct host app (PS for raster, AI for vector).
 */

/* ------------------------------------------------------------------ */
/*  External type stubs that ExtendScript lacks                       */
/* ------------------------------------------------------------------ */
declare const app: {
  selection: Array<any>;
  activeDocument: Document;
};
declare const BridgeTalk: typeof import('./BridgeTalk');
declare enum GuideOrientation { HORIZONTAL = 'h', VERTICAL = 'v' }

interface Document {
  readonly fullName: string;
}

interface Link {
  readonly filePath: string;
}

interface Image {
  readonly itemLink: Link | null;
  readonly parent: any;
}

/* ------------------------------------------------------------------ */
/*  Pure utilities – no side effects                                  */
/* ------------------------------------------------------------------ */
const isRaster = (ext: string): boolean =>
  ['.psd', '.jpg', '.jpeg', '.tif', '.tiff', '.png'].includes(ext);

const isVector = (ext: string): boolean =>
  ['.ai', '.eps', '.svg'].includes(ext);

const normalizeExt = (path: string): string =>
  File(path).name.toLowerCase().slice(-4);

/* ------------------------------------------------------------------ */
/*  Domain errors                                                     */
/* ------------------------------------------------------------------ */
class SelectionError extends Error {
  constructor(msg = 'Select a single frame that contains a linked image') {
    super(msg);
    this.name = 'SelectionError';
  }
}

class LinkError extends Error {
  constructor() {
    super('Frame does not contain a linked file');
    this.name = 'LinkError';
  }
}

class FileTypeError extends Error {
  constructor(ext: string) {
    super(`Unsupported file type “${ext}”`);
    this.name = 'FileTypeError';
  }
}

/* ------------------------------------------------------------------ */
/*  Core logic – keeps IO on the edges                                */
/* ------------------------------------------------------------------ */
function extractImage(sel: any[]): Image {
  if (sel.length !== 1) throw new SelectionError();
  const [item] = sel;

  // Case 1: user clicked the image itself
  if (item.constructor.name === 'Image') return item;

  // Case 2: user clicked the frame – grab first graphic inside
  if (item.allGraphics?.length) return item.allGraphics[0];

  throw new SelectionError();
}

function resolveTargetApp(link: Link): 'photoshop' | 'illustrator' {
  const ext = normalizeExt(link.filePath);
  if (isRaster(ext)) return 'photoshop';
  if (isVector(ext)) return 'illustrator';
  throw new FileTypeError(ext);
}

function openFileInHost(path: string, host: 'photoshop' | 'illustrator'): void {
  const bt = new BridgeTalk();
  bt.target = host;
  bt.body   = `app.open(new File("${path.replace(/\\/g, '/')}"));`;
  bt.onError = (e: any) => $.writeln(`BridgeTalk error: ${e.body}`);
  bt.send();
}

/* ------------------------------------------------------------------ */
/*  Public API – single entry point                                   */
/* ------------------------------------------------------------------ */
function main(): void {
  try {
    const img  = extractImage(app.selection);
    const link = img.itemLink ?? (() => { throw new LinkError(); })();
    const host = resolveTargetApp(link);

    openFileInHost(link.filePath, host);
  } catch (err) {
    alert(err instanceof Error ? err.message : String(err));
  }
}

/* ------------------------------------------------------------------ */
/*  Immediate invocation when InDesign runs the file                  */
/* ------------------------------------------------------------------ */
main();