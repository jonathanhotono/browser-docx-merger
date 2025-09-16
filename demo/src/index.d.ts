type MergeOptions = {
    pattern?: string | null;
    insertStart?: boolean;
    insertEnd?: boolean;
    pageBreaks?: boolean;
    mergeNumbering?: boolean;
    mergeStyles?: boolean;
    mergeFootnotes?: boolean;
    onLog?: (msg: string, level?: 'info' | 'ok' | 'warn' | 'err') => void;
};
declare function mergeDocx(inputBuffers: (ArrayBuffer | Uint8Array | Blob)[], options?: MergeOptions): Promise<Blob>;
declare function mergeDocxFromFiles(files: File[], options?: MergeOptions): Promise<Blob>;
declare function triggerDownload(url: string, name: string): void;

export { type MergeOptions, mergeDocx, mergeDocxFromFiles, triggerDownload };
