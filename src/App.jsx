import { useRef, useState, useCallback, useEffect } from "react";
import html2canvas from "html2canvas";
import Cropper from "react-easy-crop";
import { Document, Packer, Paragraph, ImageRun, PageBreak } from "docx";
import { saveAs } from "file-saver";
import "./index.css";

const emptyGroup = () => ({
  invoice: null,
  receipt: null,
  idFront: null,
  idBack: null,
});

const fields = [
  ["invoice", "发票"],
  ["receipt", "长小票"],
  ["idFront", "身份证正面"],
  ["idBack", "身份证反面"],
];

/* ─── IndexedDB 存储（容量远超 localStorage） ─── */
const DB_NAME = "photos_guobu_db";
const DB_VERSION = 1;

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains("groups")) {
        request.result.createObjectStore("groups", { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveGroups(groups) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("groups", "readwrite");
    tx.objectStore("groups").put({ id: "data", groups });
    tx.oncomplete = resolve;
    tx.onerror = reject;
  });
}

async function loadGroups() {
  try {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction("groups", "readonly");
      const req = tx.objectStore("groups").get("data");
      req.onsuccess = () => resolve(req.result?.groups || null);
      req.onerror = reject;
    });
  } catch (e) {
    return null;
  }
}

/* ─── 工具函数 ─── */
function createImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

async function fileToDataUrl(file) {
  // 先读取原始文件
  const raw = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  // 超过 2MB 的图片压缩后再存储，节省 IndexedDB 空间
  if (raw.length > 2 * 1024 * 1024) {
    try {
      const img = await createImage(raw);
      const MAX = 1600;
      let w = img.width;
      let h = img.height;
      if (Math.max(w, h) > MAX) {
        const r = MAX / Math.max(w, h);
        w = Math.round(w * r);
        h = Math.round(h * r);
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        return canvas.toDataURL("image/jpeg", 0.75);
      }
    } catch (e) { /* 压缩失败就用原图 */ }
  }
  return raw;
}

async function getImageFromDrop(e) {
  const dt = e.dataTransfer;

  if (dt.files.length > 0) return dt.files[0];

  if (dt.items) {
    for (const item of dt.items) {
      if (item.kind === "file") {
        const file = item.getAsFile();
        if (file && file.size > 0) return file;
      }
    }
  }

  const html = dt.getData("text/html");
  if (html) {
    const match = html.match(/<img[^>]+src="(data:image\/[^"]+)"/i);
    if (match?.[1]) {
      const resp = await fetch(match[1]);
      const blob = await resp.blob();
      return new File([blob], "image.png", { type: blob.type });
    }
  }

  return null;
}

async function getCroppedImage(imageSrc, pixelCrop, rotation = 0) {
  const image = await createImage(imageSrc);
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");

  const radians = (rotation * Math.PI) / 180;
  const sin = Math.abs(Math.sin(radians));
  const cos = Math.abs(Math.cos(radians));

  const rotatedWidth = image.width * cos + image.height * sin;
  const rotatedHeight = image.width * sin + image.height * cos;

  canvas.width = rotatedWidth;
  canvas.height = rotatedHeight;

  ctx.translate(rotatedWidth / 2, rotatedHeight / 2);
  ctx.rotate(radians);
  ctx.drawImage(image, -image.width / 2, -image.height / 2);

  const croppedCanvas = document.createElement("canvas");
  const croppedCtx = croppedCanvas.getContext("2d");

  croppedCanvas.width = pixelCrop.width;
  croppedCanvas.height = pixelCrop.height;

  croppedCtx.drawImage(
    canvas,
    pixelCrop.x, pixelCrop.y,
    pixelCrop.width, pixelCrop.height,
    0, 0,
    pixelCrop.width, pixelCrop.height
  );

  return croppedCanvas.toDataURL("image/jpeg", 0.85);
}

/* ─── 主组件 ─── */
export default function App() {
  const [groups, setGroups] = useState([emptyGroup()]);
  const [dbReady, setDbReady] = useState(false);
  const [cropModal, setCropModal] = useState(null);
  const [activeGroup, setActiveGroup] = useState(0);
  const [saveStatus, setSaveStatus] = useState("");
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(null);

  const [crop, setCrop] = useState({ x: 0, y: 0 });
  const [zoom, setZoom] = useState(1);
  const [rotation, setRotation] = useState(0);
  const [croppedAreaPixels, setCroppedAreaPixels] = useState(null);

  const pagesRef = useRef(null);
  const loadIdRef = useRef(0);
  const [dragOver, setDragOver] = useState(false);
  const [tipMsg, setTipMsg] = useState("");
  const exportAbortRef = useRef(false);

  // 启动时从 IndexedDB 恢复数据
  useEffect(() => {
    loadGroups().then((saved) => {
      if (saved && Array.isArray(saved) && saved.length > 0) {
        setGroups(saved.map((g) => ({
          invoice: g.invoice || null,
          receipt: g.receipt || null,
          idFront: g.idFront || null,
          idBack: g.idBack || null,
        })));
      }
    }).catch(() => {}).finally(() => setDbReady(true));
  }, []);

  // 数据变化 → 自动存入 IndexedDB
  useEffect(() => {
    if (!dbReady) return;
    const timer = setTimeout(() => {
      saveGroups(groups)
        .then(() => setSaveStatus("已自动保存"))
        .catch(() => setSaveStatus("保存失败"));
    }, 800);
    return () => clearTimeout(timer);
  }, [groups, dbReady]);

  // 全局拖拽提示
  useEffect(() => {
    const onDragEnter = () => setDragOver(true);
    const onDragLeave = (e) => {
      if (!e.relatedTarget || e.relatedTarget === document.documentElement) {
        setDragOver(false);
      }
    };
    document.addEventListener("dragenter", onDragEnter);
    document.addEventListener("dragleave", onDragLeave);
    return () => {
      document.removeEventListener("dragenter", onDragEnter);
      document.removeEventListener("dragleave", onDragLeave);
    };
  }, []);

  // 全局 Ctrl+V 粘贴：贴到当前激活组第一个空位
  useEffect(() => {
    const onPaste = async (e) => {
      if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") return;
      const items = e.clipboardData?.items;
      if (!items) return;
      for (const item of items) {
        if (item.type.startsWith("image/")) {
          e.preventDefault();
          const file = item.getAsFile();
          if (file) {
            const group = groups[activeGroup];
            if (!group) return;
            for (const [key] of fields) {
              if (!group[key]) {
                setImage(activeGroup, key, file);
                setTipMsg(`已粘贴到第 ${activeGroup + 1} 组 - ${fields.find(([k]) => k === key)[1]}`);
                setTimeout(() => setTipMsg(""), 2000);
                return;
              }
            }
          }
        }
      }
    };
    document.addEventListener("paste", onPaste);
    return () => document.removeEventListener("paste", onPaste);
  }, [groups, activeGroup]);

  /* ─── 图片操作 ─── */
  const setImage = async (groupIndex, key, file) => {
    if (!file) return;

    setSaveStatus("正在处理...");
    const id = ++loadIdRef.current;
    const url = await fileToDataUrl(file);
    if (id !== loadIdRef.current) return;

    try {
      await createImage(url);
    } catch (e) {
      setSaveStatus("该文件不是有效图片");
      return;
    }

    setGroups((prev) => {
      const copy = [...prev];
      copy[groupIndex] = { ...copy[groupIndex], [key]: url };
      return copy;
    });
  };

  const clearImage = (groupIndex, key) => {
    setGroups((prev) => {
      const copy = [...prev];
      copy[groupIndex] = { ...copy[groupIndex], [key]: null };
      return copy;
    });
  };

  const addGroup = () => {
    setGroups((prev) => {
      const next = [...prev, emptyGroup()];
      setTimeout(() => {
        const page = document.getElementById(`page-${next.length - 1}`);
        page?.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 50);
      return next;
    });
  };

  const removeGroup = (index) => {
    setGroups((prev) => prev.filter((_, i) => i !== index));
  };

  const scrollToPage = (index) => {
    const page = document.getElementById(`page-${index}`);
    if (page) page.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  /* ─── 滚动联动 ─── */
  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        const visible = entries
          .filter((e) => e.isIntersecting)
          .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
        if (visible.length > 0) {
          const id = visible[0].target.id;
          const index = parseInt(id.replace("page-", ""), 10);
          if (!isNaN(index)) setActiveGroup(index);
        }
      },
      { threshold: 0.3 }
    );

    const pages = pagesRef.current?.querySelectorAll(".a4-page");
    pages?.forEach((p) => observer.observe(p));
    return () => observer.disconnect();
  }, [groups]);

  useEffect(() => {
    const panel = document.getElementById(`group-${activeGroup}`);
    panel?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, [activeGroup]);

  /* ─── 裁剪 ─── */
  const openCrop = (groupIndex, key) => {
    const image = groups[groupIndex][key];
    if (!image) return;
    setCropModal({ groupIndex, key, image });
    setCrop({ x: 0, y: 0 });
    setZoom(1);
    setRotation(0);
    setCroppedAreaPixels(null);
  };

  const onCropComplete = useCallback((_, croppedPixels) => {
    setCroppedAreaPixels(croppedPixels);
  }, []);

  const applyCrop = async () => {
    if (!cropModal || !croppedAreaPixels) return;
    const croppedImage = await getCroppedImage(
      cropModal.image, croppedAreaPixels, rotation
    );
    setGroups((prev) => {
      const copy = [...prev];
      copy[cropModal.groupIndex] = {
        ...copy[cropModal.groupIndex],
        [cropModal.key]: croppedImage,
      };
      return copy;
    });
    setCropModal(null);
  };

  /* ─── 拖放 / 粘贴 ─── */
  const handleDropFail = () => {
    setTipMsg("此来源不支持拖拽，请在 WPS 中 Ctrl+C 复制图片，再回到网页 Ctrl+V 粘贴");
    setTimeout(() => setTipMsg(""), 4000);
  };

  const handleDrop = async (e, groupIndex, key) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOver(false);
    const file = await getImageFromDrop(e);
    if (file) {
      setImage(groupIndex, key, file);
    } else {
      handleDropFail();
    }
  };

  /* ─── 导出 Word（批量渲染 + 内存优化 + 进度条） ─── */
  const exportWord = async () => {
    const pages = pagesRef.current.querySelectorAll(".a4-page");
    if (pages.length === 0) { alert("没有可导出的页面"); return; }

    exportAbortRef.current = false;
    setExporting(true);
    setExportProgress({ current: 0, total: pages.length });

    // 先收集并清除样式
    const originals = [];
    pages.forEach((p) => {
      originals.push({ padding: p.style.padding, boxShadow: p.style.boxShadow });
      p.style.padding = "0";
      p.style.boxShadow = "none";
    });

    const children = [];

    try {
      for (let i = 0; i < pages.length; i++) {
        if (exportAbortRef.current) break;

        try {
          const canvas = await html2canvas(pages[i], {
            scale: 2,
            backgroundColor: "#ffffff",
            ignoreElements: (el) => el.hasAttribute("data-html2canvas-ignore"),
          });

          // JPEG 质量 0.95，与 PNG 视觉差异极小
          const blob = await new Promise((resolve) =>
            canvas.toBlob(resolve, "image/jpeg", 0.95)
          );

          const buffer = await blob.arrayBuffer();

          // A4 页面在 docx 中精确宽度为 793.7px（96DPI），取 793 防止溢出导致空白页
          const aspectRatio = canvas.width / canvas.height;
          const imgWidth = 793;
          const imgHeight = Math.round(imgWidth / aspectRatio);

          children.push(
            new Paragraph({
              spacing: { before: 0, after: 0 },
              children: [
                new ImageRun({
                  data: buffer,
                  transformation: { width: imgWidth, height: imgHeight },
                }),
              ],
            })
          );

          // 释放 canvas 内存
          canvas.width = 0;
          canvas.height = 0;
        } catch (err) {
          console.warn(`第 ${i + 1} 页渲染失败，已跳过`, err);
        }

        setExportProgress({ current: i + 1, total: pages.length });
        // 每 5 页让出主线程，避免 UI 卡死
        if (i % 5 === 4) await new Promise((r) => setTimeout(r, 0));
      }
    } finally {
      pages.forEach((p, i) => {
        p.style.padding = originals[i]?.padding ?? "";
        p.style.boxShadow = originals[i]?.boxShadow ?? "";
      });
    }

    if (children.length === 0) {
      alert("没有可导出的页面");
      setExporting(false);
      setExportProgress(null);
      return;
    }

    // 组装文档
    const docChildren = [];
    for (let i = 0; i < children.length; i++) {
      docChildren.push(children[i]);
      if (i < children.length - 1) {
        docChildren.push(
          new Paragraph({
            spacing: { before: 0, after: 0 },
            children: [new PageBreak()],
          })
        );
      }
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 0, right: 0, bottom: 0, left: 0 },
          },
        },
        children: docChildren,
      }],
    });

    const docBlob = await Packer.toBlob(doc);
    saveAs(docBlob, "资料汇总.docx");

    setExporting(false);
    setExportProgress(null);
  };

  /* ─── 渲染 ─── */
  return (
    <div className="app">
      <aside className="toolbar" data-html2canvas-ignore="true">
        <div className="toolbar-header">
          <h1>批量资料排版</h1>
          <p>支持上传、拖拽、清除、裁剪微调，最后一键生成多页 Word。</p>

          <button onClick={addGroup} disabled={exporting}>新增一组</button>
          <button onClick={exportWord} disabled={exporting}>
            {exporting ? "导出中..." : "生成 Word"}
          </button>

          {exportProgress && (
            <div className="export-progress">
              <div className="progress-bar">
                <div
                  className="progress-fill"
                  style={{ width: `${(exportProgress.current / exportProgress.total) * 100}%` }}
                />
              </div>
              <span>{exportProgress.current} / {exportProgress.total} 页</span>
              <button
                className="cancel-export-btn"
                onClick={() => { exportAbortRef.current = true; }}
              >
                取消导出
              </button>
            </div>
          )}

          {saveStatus && <div className="save-status">{saveStatus}</div>}
          <button className="clear-data-btn" onClick={() => {
            if (confirm("确定要清除所有本地保存的数据吗？此操作不可恢复。")) {
              openDB().then((db) => {
                const tx = db.transaction("groups", "readwrite");
                tx.objectStore("groups").delete("data");
              }).catch(() => {});
              setGroups([emptyGroup()]);
              setSaveStatus("数据已清除");
            }
          }}>清除保存数据</button>
        </div>

        <div className="toolbar-groups">
          {groups.map((group, groupIndex) => (
            <div
              id={`group-${groupIndex}`}
              className={`group-panel${groupIndex === activeGroup ? " active" : ""}`}
              key={groupIndex}
              onClick={() => scrollToPage(groupIndex)}
            >
              <div className="group-title">
                <strong>第 {groupIndex + 1} 组</strong>
                {groups.length > 1 && (
                  <button
                    className="delete-btn"
                    onClick={() => removeGroup(groupIndex)}
                  >
                    删除本组
                  </button>
                )}
              </div>

              {fields.map(([key, label]) => (
                <div className="upload-row" key={key}>
                  <label
                    className={`upload-card ${group[key] ? "has-file" : ""}`}
                    tabIndex={0}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => handleDrop(e, groupIndex, key)}
                    onPaste={async (e) => {
                      const items = e.clipboardData?.items;
                      if (!items) return;
                      for (const item of items) {
                        if (item.type.startsWith("image/")) {
                          e.preventDefault();
                          const file = item.getAsFile();
                          if (file) {
                            setImage(groupIndex, key, file);
                            return;
                          }
                        }
                      }
                    }}
                  >
                    <strong>{label}</strong>
                    <span>
                      {group[key] ? "已上传，可重新选择" : "点击选择、拖拽或 Ctrl+V 粘贴"}
                    </span>
                    <input
                      type="file"
                      accept="image/*"
                      onChange={(e) => setImage(groupIndex, key, e.target.files[0])}
                    />
                  </label>

                  <div className="small-actions">
                    {group[key] && (
                      <>
                        <button className="mini-btn crop-btn" onClick={() => openCrop(groupIndex, key)}>
                          裁剪
                        </button>
                        <button className="mini-btn clear-btn" onClick={() => clearImage(groupIndex, key)}>
                          清除
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))}
            </div>
          ))}
        </div>
      </aside>

      {dragOver && <div className="drag-overlay">拖放图片到对应区域，或使用 Ctrl+V 粘贴</div>}
      {tipMsg && <div className="tip-toast">{tipMsg}</div>}

      <main className="preview-wrap" ref={pagesRef}>
        {groups.map((group, index) => (
          <div className="a4-page" key={index} id={`page-${index}`}>
            <div className="page-label" data-html2canvas-ignore="true">
              第 {index + 1} 页
            </div>

            <section
              className="invoice-area"
              style={group.invoice ? undefined : { border: "none" }}
              onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
              onDrop={(e) => handleDrop(e, index, "invoice")}
            >
              {group.invoice ? (
                <img src={group.invoice} className="invoice-img" draggable={false} />
              ) : (
                <div className="placeholder" data-html2canvas-ignore="true">发票</div>
              )}
            </section>

            <section className="bottom-area">
              <div
                className="receipt-area"
                style={group.receipt ? undefined : { border: "none" }}
                onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                onDrop={(e) => handleDrop(e, index, "receipt")}
              >
                {group.receipt ? (
                  <div className="receipt-split">
                    <div className="receipt-half">
                      <img src={group.receipt} className="receipt-first" draggable={false} />
                    </div>
                    <div className="receipt-half">
                      <img src={group.receipt} className="receipt-second" draggable={false} />
                    </div>
                  </div>
                ) : (
                  <div className="placeholder" data-html2canvas-ignore="true">长小票</div>
                )}
              </div>

              <div className="id-area">
                <div
                  className="id-box"
                  style={group.idFront ? undefined : { border: "none" }}
                  onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                  onDrop={(e) => handleDrop(e, index, "idFront")}
                >
                  {group.idFront ? (
                    <img src={group.idFront} draggable={false} />
                  ) : (
                    <div className="placeholder" data-html2canvas-ignore="true">身份证正面</div>
                  )}
                </div>
                <div
                  className="id-box"
                  style={group.idBack ? undefined : { border: "none" }}
                  onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                  onDrop={(e) => handleDrop(e, index, "idBack")}
                >
                  {group.idBack ? (
                    <img src={group.idBack} draggable={false} />
                  ) : (
                    <div className="placeholder" data-html2canvas-ignore="true">身份证反面</div>
                  )}
                </div>
              </div>
            </section>
          </div>
        ))}
      </main>

      {cropModal && (
        <div className="crop-modal" data-html2canvas-ignore="true">
          <div className="crop-panel">
            <h2>裁剪图片</h2>
            <div className="crop-box">
              <Cropper
                image={cropModal.image}
                crop={crop}
                zoom={zoom}
                rotation={rotation}
                aspect={cropModal.key === "receipt" ? 1 / 2 : 16 / 10}
                onCropChange={setCrop}
                onZoomChange={setZoom}
                onRotationChange={setRotation}
                onCropComplete={onCropComplete}
              />
            </div>

            <div className="range-row">
              <label>缩放</label>
              <input
                type="range"
                min={1} max={4} step={0.01}
                value={zoom}
                onChange={(e) => setZoom(Number(e.target.value))}
              />
            </div>

            <div className="rotate-actions">
              <button onClick={() => setRotation((prev) => prev - 90)}>向左旋转90°</button>
              <button onClick={() => setRotation((prev) => prev + 90)}>向右旋转90°</button>
            </div>

            <div className="crop-actions">
              <button onClick={applyCrop}>应用裁剪</button>
              <button className="gray-btn" onClick={() => setCropModal(null)}>取消</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
