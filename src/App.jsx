import { useRef, useState, useCallback } from "react";
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

function createImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
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
    pixelCrop.x,
    pixelCrop.y,
    pixelCrop.width,
    pixelCrop.height,
    0,
    0,
    pixelCrop.width,
    pixelCrop.height
  );

  return croppedCanvas.toDataURL("image/jpeg", 0.95);
}

export default function App() {
  const [groups, setGroups] = useState([emptyGroup()]);
  const [cropModal, setCropModal] = useState(null);

  const [crop, setCrop] = useState({ x: 0, y: 0 });
  const [zoom, setZoom] = useState(1);
  const [rotation, setRotation] = useState(0);
  const [croppedAreaPixels, setCroppedAreaPixels] = useState(null);

  const pagesRef = useRef(null);

  const setImage = (groupIndex, key, file) => {
    if (!file || !file.type.startsWith("image/")) return;

    const url = URL.createObjectURL(file);

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
    setGroups((prev) => [...prev, emptyGroup()]);
  };

  const removeGroup = (index) => {
    setGroups((prev) => prev.filter((_, i) => i !== index));
  };

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
      cropModal.image,
      croppedAreaPixels,
      rotation
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

  const exportWord = async () => {
    const pages = pagesRef.current.querySelectorAll(".a4-page");
    const children = [];

    for (let i = 0; i < pages.length; i++) {
      try {
        const canvas = await html2canvas(pages[i], {
          scale: 2,
          backgroundColor: "#ffffff",
          ignoreElements: (element) =>
            element.hasAttribute("data-html2canvas-ignore"),
        });

        const blob = await new Promise((resolve) =>
          canvas.toBlob(resolve, "image/png")
        );

        const buffer = await blob.arrayBuffer();

        children.push(
          new Paragraph({
            children: [
              new ImageRun({
                data: buffer,
                transformation: {
                  width: 595,
                  height: 842,
                },
              }),
            ],
          })
        );
      } catch (err) {
        console.warn(`第 ${i + 1} 页渲染失败，已跳过`, err);
      }
    }

    if (children.length === 0) {
      alert("没有可导出的页面");
      return;
    }

    const docChildren = [];
    for (let i = 0; i < children.length; i++) {
      docChildren.push(children[i]);
      if (i !== children.length - 1) {
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));
      }
    }

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 0,
                right: 0,
                bottom: 0,
                left: 0,
              },
            },
          },
          children: docChildren,
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "资料汇总.docx");
  };

  return (
    <div className="app">
      <aside className="toolbar" data-html2canvas-ignore="true">
        <div className="toolbar-header">
          <h1>批量资料排版</h1>
          <p>支持上传、拖拽、清除、裁剪微调，最后一键生成多页 Word。</p>

          <button onClick={addGroup}>新增一组</button>
          <button onClick={exportWord}>生成 Word</button>
        </div>

        <div className="toolbar-groups">
          {groups.map((group, groupIndex) => (
          <div className="group-panel" key={groupIndex}>
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
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    setImage(groupIndex, key, e.dataTransfer.files[0]);
                  }}
                >
                  <strong>{label}</strong>
                  <span>
                    {group[key] ? "已上传，可重新选择" : "点击选择或拖拽图片"}
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
                      <button
                        className="mini-btn crop-btn"
                        onClick={() => openCrop(groupIndex, key)}
                      >
                        裁剪
                      </button>

                      <button
                        className="mini-btn clear-btn"
                        onClick={() => clearImage(groupIndex, key)}
                      >
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

      <main className="preview-wrap" ref={pagesRef}>
        {groups.map((group, index) => (
          <div className="a4-page" key={index}>
            <div className="page-label" data-html2canvas-ignore="true">
              第 {index + 1} 页
            </div>

            <section className="invoice-area" style={group.invoice ? undefined : { border: "none" }}>
              {group.invoice ? (
                <img src={group.invoice} className="invoice-img" />
              ) : (
                <div className="placeholder" data-html2canvas-ignore="true">发票</div>
              )}
            </section>

            <section className="bottom-area">
              <div className="receipt-area" style={group.receipt ? undefined : { border: "none" }}>
                {group.receipt ? (
                  <div className="receipt-split">
                    <div className="receipt-half">
                      <img src={group.receipt} className="receipt-first" />
                    </div>

                    <div className="receipt-half">
                      <img src={group.receipt} className="receipt-second" />
                    </div>
                  </div>
                ) : (
                  <div className="placeholder" data-html2canvas-ignore="true">长小票</div>
                )}
              </div>

              <div className="id-area">
                <div className="id-box" style={group.idFront ? undefined : { border: "none" }}>
                  {group.idFront ? (
                    <img src={group.idFront} />
                  ) : (
                    <div className="placeholder" data-html2canvas-ignore="true">身份证正面</div>
                  )}
                </div>

                <div className="id-box" style={group.idBack ? undefined : { border: "none" }}>
                  {group.idBack ? (
                    <img src={group.idBack} />
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
                min={1}
                max={4}
                step={0.01}
                value={zoom}
                onChange={(e) => setZoom(Number(e.target.value))}
              />
            </div>

            <div className="rotate-actions">
              <button onClick={() => setRotation((prev) => prev - 90)}>
                向左旋转90°
              </button>

              <button onClick={() => setRotation((prev) => prev + 90)}>
                向右旋转90°
              </button>
            </div>


            <div className="crop-actions">
              <button onClick={applyCrop}>应用裁剪</button>
              <button className="gray-btn" onClick={() => setCropModal(null)}>
                取消
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
