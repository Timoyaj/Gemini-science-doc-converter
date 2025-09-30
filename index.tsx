/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { GoogleGenAI } from "@google/genai";
import {
  useState,
  useCallback,
  useRef,
  useEffect,
} from "https://esm.sh/preact/hooks";

// These are globals from the imported scripts in index.html
declare var h: any;
declare var html: any;
declare var pdfjsLib: any;
declare var docx: any;

interface SavedConversion {
  id: number;
  fileName: string;
  docxBase64: string;
  preview: string;
  createdAt: string;
}

const App = () => {
  const [file, setFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isPreviewLoading, setIsPreviewLoading] = useState(false);
  const [docxUrl, setDocxUrl] = useState<string | null>(null);
  const [generatedDocxBlob, setGeneratedDocxBlob] = useState<Blob | null>(null);
  const [isCurrentSaved, setIsCurrentSaved] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [savedConversions, setSavedConversions] = useState<SavedConversion[]>([]);
  const [useOcr, setUseOcr] = useState(false);
  const dropzoneRef = useRef<HTMLLabelElement>(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem('savedConversions');
      if (saved) {
        setSavedConversions(JSON.parse(saved));
      }
    } catch (e) {
      console.error("Failed to load saved conversions from localStorage", e);
      setError("Could not load previously saved documents.");
    }
  }, []);

  const handleFileChange = async (selectedFile: File | null) => {
    if (!selectedFile) return;

    setError(null);
    setDocxUrl(null);
    setGeneratedDocxBlob(null);
    setIsCurrentSaved(false);
    setFile(selectedFile);
    setPreview(null);
    setIsPreviewLoading(true);
    setUseOcr(false);

    const fileType = selectedFile.type;
    if (fileType.startsWith("image/")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        setPreview(e.target?.result as string);
        setIsPreviewLoading(false);
      };
      reader.onerror = () => {
        setError("Could not read the image file.");
        setIsPreviewLoading(false);
      }
      reader.readAsDataURL(selectedFile);
    } else if (fileType === "application/pdf") {
      const fileReader = new FileReader();
      fileReader.onload = async (e) => {
        try {
          const typedarray = new Uint8Array(e.target?.result as ArrayBuffer);
          const pdf = await pdfjsLib.getDocument(typedarray).promise;
          const page = await pdf.getPage(1);
          const viewport = page.getViewport({ scale: 1.5 });
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) {
            setError("Could not create canvas context for PDF preview.");
            return;
          }
          canvas.height = viewport.height;
          canvas.width = viewport.width;
          await page.render({ canvasContext: context, viewport: viewport }).promise;
          setPreview(canvas.toDataURL());
        } catch (err: any) {
          setError(`Failed to render PDF preview: ${err.message}`);
          console.error(err);
        } finally {
          setIsPreviewLoading(false);
        }
      };
      fileReader.onerror = () => {
        setError("Could not read the PDF file. It might be corrupted.");
        setIsPreviewLoading(false);
      };
      fileReader.readAsArrayBuffer(selectedFile);
    } else {
      setError("Unsupported file type. Please upload an image or a PDF.");
      setFile(null);
      setPreview(null);
      setIsPreviewLoading(false);
    }
  };
  
  const onDragOver = useCallback((e: DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e: DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const onDrop = useCallback((e: DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer?.files?.length) {
      handleFileChange(e.dataTransfer.files[0]);
    }
  }, []);


  const handleConvert = async () => {
    if (!file) return;
    setIsLoading(true);
    setError(null);
    setDocxUrl(null);
    setGeneratedDocxBlob(null);
    setIsCurrentSaved(false);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

      const fileToParts = async (file: File) => {
        const base64 = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve((reader.result as string).split(',')[1]);
            reader.onerror = (error) => reject(error);
        });
        return {
            inlineData: {
                mimeType: file.type,
                data: base64
            }
        };
      };

      const imageParts = [];
      if (file.type === 'application/pdf') {
          const fileReader = new FileReader();
          const buffer = await new Promise<ArrayBuffer>((resolve) => {
              fileReader.onload = (e) => resolve(e.target?.result as ArrayBuffer);
              fileReader.readAsArrayBuffer(file);
          });
          const typedarray = new Uint8Array(buffer);
          const pdf = await pdfjsLib.getDocument(typedarray).promise;
          for (let i = 1; i <= pdf.numPages; i++) {
              const page = await pdf.getPage(i);
              const viewport = page.getViewport({ scale: 2.0 }); // Higher scale for better quality
              const canvas = document.createElement("canvas");
              const context = canvas.getContext("2d");
              canvas.height = viewport.height;
              canvas.width = viewport.width;
              await page.render({ canvasContext: context, viewport }).promise;
              const blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, 'image/png'));
              if(blob) {
                const pageFile = new File([blob], `page_${i}.png`, {type: 'image/png'});
                imageParts.push(await fileToParts(pageFile));
              }
          }
      } else {
          imageParts.push(await fileToParts(file));
      }

      let prompt = `You are an expert at document analysis and conversion. Your task is to accurately extract all content from the provided image(s), including all text and mathematical equations.
- Preserve the original structure, paragraphs, and formatting as closely as possible.
- Identify ALL mathematical equations, both inline and display.
- Convert every mathematical equation and formula into its corresponding LaTeX format.
- For INLINE equations (within a line of text), wrap the LaTeX in single dollar signs: $...$.
- For DISPLAY equations (on their own line), wrap the LaTeX in double dollar signs: $$...$$.
- Ensure the output is a single, continuous block of text that correctly intersperses regular text and LaTeX-formatted equations in their original order.
Do not add any commentary, explanations, or markdown formatting like headers or lists. Only output the extracted content.`;

      if (file.type.startsWith("image/") && useOcr) {
        prompt = `You are an expert at document analysis and conversion, equipped with advanced Optical Character Recognition (OCR) capabilities. The following image may be low-resolution, have poor lighting, or a complex background.
- Apply your most powerful OCR techniques to accurately extract all content, including all text and mathematical equations.
- Preserve the original structure, paragraphs, and formatting as closely as possible.
- Identify ALL mathematical equations, both inline and display.
- Convert every mathematical equation and formula into its corresponding LaTeX format.
- For INLINE equations (within a line of text), wrap the LaTeX in single dollar signs: $...$.
- For DISPLAY equations (on their own line), wrap the LaTeX in double dollar signs: $$...$$.
- Ensure the output is a single, continuous block of text that correctly intersperses regular text and LaTeX-formatted equations in their original order.
Do not add any commentary, explanations, or markdown formatting like headers or lists. Only output the extracted content.`;
      }

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: { parts: [...imageParts, {text: prompt}] },
      });

      const text = response.text;
      await generateDocx(text);
    } catch (err: any) {
      setError(`An error occurred during conversion: ${err.message}`);
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };
  
  const preprocessLatexForWord = (latex: string): string => {
    let processed = latex.trim();

    // 1. Handle multi-line equations by replacing \\ with @.
    // Word's equation editor uses @ for new lines in equation arrays.
    processed = processed.replace(/\\\\/g, '@');

    // 2. Delimiters: \left( becomes (, \right) becomes ), etc.
    processed = processed.replace(/\\left\(/g, '(');
    processed = processed.replace(/\\right\)/g, ')');
    processed = processed.replace(/\\left\[/g, '[');
    processed = processed.replace(/\\right\]/g, ']');
    processed = processed.replace(/\\left\{/g, '{');
    processed = processed.replace(/\\right\}/g, '}');

    // 3. Standardize function calls, e.g., \sin{x} to \sin(x)
    processed = processed.replace(/\\(sin|cos|tan|lim|log|ln|det|exp)\s*\{(.*?)\}/g, '\\$1($2)');

    // 4. Fractions and binomials
    processed = processed.replace(/\\frac{(.*?)}{(.*?)}/g, '($1)/($2)');
    processed = processed.replace(/\\binom{(.*?)}{(.*?)}/g, '($1\atop $2)');

    // 5. Roots
    processed = processed.replace(/\\sqrt\[(.*?)\]{(.*?)}/g, '\\root($1&$2)');
    processed = processed.replace(/\\sqrt{(.*?)}/g, '\\sqrt($1)');

    // 6. Accents and decorations
    processed = processed.replace(/\\(vec|hat|bar|dot|ddot|tilde){(.*?)}/g, '\\$1($2)');

    // 7. Handle matrix environments (pmatrix, bmatrix, vmatrix, Vmatrix, matrix)
    processed = processed.replace(/\\begin{(p|b|v|V|)matrix}([\s\S]*?)\\end{(p|b|v|)matrix}/g, (match, type, content) => {
        // In Word, @ is newline, & is column separator. We already replaced \\ with @
        const rows = content.trim().replace(/\\\\/g, '@');
        switch (type) {
            case 'p': return `(${rows})`; 
            case 'b': return `[${rows}]`;
            case 'v': return `|${rows}|`;
            case 'V': return `‖${rows}‖`;
            case '': return `\\matrix(${rows})`;
            default: return `\\matrix(${rows})`;
        }
    });
    
    // 8. General subscripts and superscripts for commands and variables.
    // e.g., \sum_{i=0}^{n} becomes \sum_(i=0)^(n)
    // e.g., x_{i} becomes x_(i)
    processed = processed.replace(/([a-zA-Z\\]+)_\{(.*?)\}\^\{(.*?)\}/g, '$1_($2)^($3)');
    processed = processed.replace(/([a-zA-Z\\]+)_\{(.*?)\}/g, '$1_($2)');
    processed = processed.replace(/([a-zA-Z\\]+)\^\{(.*?)\}/g, '$1^($2)');

    // 9. Remove parentheses around single-character superscripts and subscripts for cleaner conversion.
    // e.g., x_(i) -> x_i and v^(2) -> v^2
    processed = processed.replace(/([_^])\(([^()]{1})\)/g, '$1$2');

    return processed;
  };

  const generateDocx = async (text: string) => {
    const paragraphs: any[] = [];
    const documentParts = text.split(/(\$\$[\s\S]*?\$\$)/g);

    for (const part of documentParts) {
        if (!part) continue;

        if (part.startsWith('$$') && part.endsWith('$$')) {
            const math = part.slice(2, -2).trim();
            if (math) {
                const preprocessedMath = preprocessLatexForWord(math);
                const mathParagraph = new docx.Paragraph({
                    children: [docx.Math.run(preprocessedMath)],
                    alignment: docx.AlignmentType.CENTER,
                });
                paragraphs.push(mathParagraph);
            }
        } else {
            const textParagraphs = part.split('\n');
            for (const paraText of textParagraphs) {
                if (!paraText.trim() && paragraphs.length > 0) continue;

                const inlineParts = paraText.split(/(\$[\s\S]*?\$)/g);
                const paragraphChildren: any[] = [];

                for (const inlinePart of inlineParts) {
                    if (!inlinePart) continue;

                    if (inlinePart.startsWith('$') && inlinePart.endsWith('$')) {
                        const math = inlinePart.slice(1, -1).trim();
                        if (math) {
                            const preprocessedMath = preprocessLatexForWord(math);
                            paragraphChildren.push(docx.Math.run(preprocessedMath));
                        }
                    } else {
                        if (inlinePart) {
                           paragraphChildren.push(new docx.TextRun(inlinePart));
                        }
                    }
                }
                if (paragraphChildren.length > 0) {
                   paragraphs.push(new docx.Paragraph({ children: paragraphChildren }));
                }
            }
        }
    }

    const doc = new docx.Document({
      sections: [{
        children: paragraphs.length > 0 ? paragraphs : [new docx.Paragraph({children:[]})],
      }],
    });

    const blob = await docx.Packer.toBlob(doc);
    setDocxUrl(URL.createObjectURL(blob));
    setGeneratedDocxBlob(blob);
  };

  const blobToBase64 = (blob: Blob): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend = () => {
        resolve(reader.result as string);
      };
      reader.onerror = (error) => reject(error);
    });
  };

  const handleSave = async () => {
    if (!generatedDocxBlob || !file || !preview) return;
    
    try {
      const docxBase64 = await blobToBase64(generatedDocxBlob);
      const newConversion: SavedConversion = {
        id: Date.now(),
        fileName: file.name,
        docxBase64,
        preview,
        createdAt: new Date().toISOString(),
      };
      
      const updatedConversions = [newConversion, ...savedConversions];
      setSavedConversions(updatedConversions);
      localStorage.setItem('savedConversions', JSON.stringify(updatedConversions));
      setIsCurrentSaved(true);
    } catch (e) {
      console.error("Failed to save conversion", e);
      setError("Could not save the document. Please try again.");
    }
  };

  const handleDelete = (idToDelete: number) => {
    const updatedConversions = savedConversions.filter(c => c.id !== idToDelete);
    setSavedConversions(updatedConversions);
    localStorage.setItem('savedConversions', JSON.stringify(updatedConversions));
  };
  
  const handleDownloadSaved = (item: SavedConversion) => {
    const byteString = atob(item.docxBase64.split(',')[1]);
    const mimeString = item.docxBase64.split(',')[0].split(':')[1].split(';')[0];
    const ab = new ArrayBuffer(byteString.length);
    const ia = new Uint8Array(ab);
    for (let i = 0; i < byteString.length; i++) {
        ia[i] = byteString.charCodeAt(i);
    }
    const blob = new Blob([ab], {type: mimeString});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = item.fileName.replace(/\.[^/.]+$/, "") + ".docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };


  return html`
    <div class="container">
      <header>
        <h1>DocuMath Converter</h1>
        <p>Convert images & PDFs with math to editable Word documents.</p>
      </header>

      <label
        for="file-input"
        class=${`dropzone ${isDragging ? 'dragging' : ''}`}
        ref=${dropzoneRef}
        onDragOver=${onDragOver}
        onDragLeave=${onDragLeave}
        onDrop=${onDrop}
      >
        <svg class="dropzone-icon" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
          <path stroke-linecap="round" stroke-linejoin="round" d="M12 16.5V9.75m0 0l-3.75 3.75M12 9.75l3.75 3.75M3 17.25V8.25c0-1.12.93-2.25 2.25-2.25h13.5c1.24 0 2.25 1.13 2.25 2.25v9c0 1.12-.93 2.25-2.25 2.25H5.25c-1.24 0-2.25-1.13-2.25-2.25z" />
        </svg>
        <p>
            <strong>Click to upload</strong> or drag and drop<br />
            Supports: Images (PNG, JPG) and PDF
        </p>
      </label>
      <input id="file-input" type="file" accept="image/*,application/pdf" onChange=${(e: any) => handleFileChange(e.target.files[0])} />

      ${error && html`<div class="error-message">${error}</div>`}

      ${file &&
      html`
        <div class="file-info">
          ${isPreviewLoading && html`<div class="preview-loader"><div class="loader"></div></div>`}
          ${!isPreviewLoading && preview &&
          html`<img src=${preview} alt="File preview" class="preview" />`}
          <div class="file-details">
            <p>${file.name}</p>
          </div>
        </div>
      `}
      
      ${file && file.type.startsWith("image/") && !docxUrl && html`
        <div class="options-container">
            <label class="ocr-toggle">
                <input type="checkbox" checked=${useOcr} onChange=${(e: any) => setUseOcr(e.target.checked)} />
                <span>Enable enhanced OCR for low-quality images</span>
            </label>
        </div>
      `}

      ${!file || docxUrl ? '' : html`
        <button class="btn" onClick=${handleConvert} disabled=${!file || isLoading}>
          ${isLoading && html`<div class="loader"></div>`}
          ${isLoading ? "Converting..." : "Convert to Word"}
        </button>
      `}

      ${docxUrl &&
      html`
        <div class="result-actions">
            <a href=${docxUrl} download=${file?.name.replace(/\.[^/.]+$/, "") + ".docx" ?? "document.docx"} class="btn btn-success">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
                Download .docx
            </a>
            <button class="btn btn-secondary" onClick=${handleSave} disabled=${isCurrentSaved}>
                ${isCurrentSaved ? 'Saved!' : 'Save Conversion'}
            </button>
        </div>
      `}
    </div>

    ${savedConversions.length > 0 && html`
        <div class="container saved-section">
            <h2>Saved Conversions</h2>
            <div class="saved-list">
                ${savedConversions.map(item => html`
                    <div class="saved-item" key=${item.id}>
                        <img src=${item.preview} alt="Preview of ${item.fileName}" class="saved-item-preview" />
                        <div class="saved-item-details">
                            <p class="saved-item-filename">${item.fileName}</p>
                            <p class="saved-item-date">
                                Saved on ${new Intl.DateTimeFormat('en-US', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(item.createdAt))}
                            </p>
                        </div>
                        <div class="saved-item-actions">
                            <button class="btn-icon" title="Download" onClick=${() => handleDownloadSaved(item)}>
                                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
                            </button>
                             <button class="btn-icon btn-danger" title="Delete" onClick=${() => handleDelete(item.id)}>
                                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/></svg>
                            </button>
                        </div>
                    </div>
                `)}
            </div>
        </div>
    `}
  `;
};

export default App;