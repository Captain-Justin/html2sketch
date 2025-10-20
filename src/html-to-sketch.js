import JSZip from 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm';

const SKETCH_APP = 'com.bohemiancoding.sketch3';
const SKETCH_VERSION = 124;

export class HtmlToSketchConverter {
  async convertDocument(document, options = {}) {
    const { viewportWidth, viewportHeight, filename = 'export.sketch' } = options;
    const rootRect = document.body.getBoundingClientRect();
    const width = viewportWidth || Math.round(rootRect.width);
    const height = viewportHeight || Math.round(rootRect.height);

    const pageId = this._id();
    const artboardId = this._id();

    const layers = this._collectLayers(document, rootRect);

    const artboard = {
      _class: 'artboard',
      do_objectID: artboardId,
      name: 'HTML Preview',
      booleanOperation: -1,
      frame: this._frame(0, 0, width, height),
      clippingMaskMode: 0,
      exportOptions: this._emptyExportOptions(),
      hasBackgroundColor: false,
      includeBackgroundColorInExport: true,
      includeInCloudUpload: true,
      isFlowHome: false,
      layerListExpandedType: 2,
      layers,
      nameIsFixed: true,
      resizingConstraint: 63,
      resizingType: 0,
      rotation: 0,
      shouldBreakMaskChain: true,
      style: this._defaultStyle(),
      verticalRulerData: this._rulerData(),
      horizontalRulerData: this._rulerData()
    };

    const page = {
      _class: 'page',
      do_objectID: pageId,
      booleanOperation: -1,
      clippingMaskMode: 0,
      exportOptions: this._emptyExportOptions(),
      hasClickThrough: false,
      horizontalRulerData: this._rulerData(),
      isLocked: false,
      layerListExpandedType: 0,
      layers: [artboard],
      name: 'Generated Page',
      resizingConstraint: 63,
      resizingType: 0,
      rotation: 0,
      shouldBreakMaskChain: false,
      verticalRulerData: this._rulerData()
    };

    const documentJson = {
      _class: 'document',
      do_objectID: this._id(),
      assets: { _class: 'assetCollection', colors: [], gradients: [], images: [] },
      colorSpace: 0,
      currentPageIndex: 0,
      enableLayerInteraction: true,
      enableSliceInteraction: true,
      foreignLayerStyles: [],
      foreignSwatches: [],
      foreignSymbols: [],
      foreignTextStyles: [],
      layerStyles: { _class: 'sharedStyleContainer', objects: [] },
      layerTextStyles: { _class: 'sharedTextStyleContainer', objects: [] },
      pages: [
        {
          _class: 'MSJSONFileReference',
          _ref_class: 'MSImmutablePage',
          _ref: `pages/${pageId}.json`
        }
      ]
    };

    const metaJson = {
      commit: '0000000000000000000000000000000000000000',
      app: SKETCH_APP,
      build: '99999',
      version: SKETCH_VERSION,
      compatibilityVersion: SKETCH_VERSION,
      appVersion: '95',
      variant: 'NONAPPSTORE',
      autosaved: 0,
      created: {
        app: SKETCH_APP,
        build: '99999',
        version: SKETCH_VERSION
      },
      saveHistory: ['NONAPPSTORE:95']
    };

    const userJson = {
      document: {
        pageListCollapsed: 0,
        pageListHeight: 85
      },
      pages: {}
    };

    const zip = new JSZip();
    zip.file('document.json', JSON.stringify(documentJson, null, 2));
    zip.file('meta.json', JSON.stringify(metaJson, null, 2));
    zip.file('user.json', JSON.stringify(userJson, null, 2));
    zip.folder('pages').file(`${pageId}.json`, JSON.stringify(page, null, 2));

    const blob = await zip.generateAsync({ type: 'blob' });
    return { blob, filename };
  }

  _collectLayers(document, rootRect) {
    const win = document.defaultView;
    const layers = [];
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          return node.textContent.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
        }

        if (node.nodeType === Node.ELEMENT_NODE) {
          const tag = node.tagName.toLowerCase();
          if (['script', 'style', 'meta', 'link'].includes(tag)) {
            return NodeFilter.FILTER_REJECT;
          }
          return NodeFilter.FILTER_ACCEPT;
        }

        return NodeFilter.FILTER_REJECT;
      }
    });

    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node.nodeType === Node.TEXT_NODE) {
        const layer = this._textLayer(node, rootRect, win);
        if (layer) {
          layers.push(layer);
        }
      } else if (node.nodeType === Node.ELEMENT_NODE && node !== document.body) {
        const layer = this._shapeLayer(node, rootRect, win);
        if (layer) {
          layers.push(layer);
        }
      }
    }

    return layers;
  }

  _shapeLayer(element, rootRect, win) {
    const rect = element.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return null;
    }

    const style = win.getComputedStyle(element);
    const layerId = this._id();
    const fillColor = this._parseColor(style.backgroundColor) || this._parseColor('rgba(255,255,255,0)');
    const borderColor = this._parseColor(style.borderColor || 'rgba(0,0,0,0)');

    return {
      _class: 'shapeGroup',
      do_objectID: layerId,
      name: `${element.tagName.toLowerCase()}#${element.className || element.id || layerId.substring(0, 5)}`,
      booleanOperation: -1,
      clippingMaskMode: 0,
      exportOptions: this._emptyExportOptions(),
      frame: this._frame(
        rect.left - rootRect.left,
        rect.top - rootRect.top,
        rect.width,
        rect.height
      ),
      hasClippingMask: false,
      isFixedToViewport: false,
      isFlippedHorizontal: false,
      isFlippedVertical: false,
      isLocked: false,
      isVisible: true,
      layerListExpandedType: 0,
      nameIsFixed: false,
      resizingConstraint: 63,
      resizingType: 0,
      rotation: 0,
      shouldBreakMaskChain: false,
      style: {
        _class: 'style',
        borders: borderColor && borderColor.alpha > 0 ? [
          {
            _class: 'border',
            isEnabled: true,
            color: borderColor,
            fillType: 0,
            position: 1,
            thickness: parseFloat(style.borderWidth) || 1
          }
        ] : [],
        fills: [
          {
            _class: 'fill',
            isEnabled: fillColor.alpha > 0,
            color: fillColor,
            fillType: 0,
            noiseIndex: 0,
            noiseIntensity: 0,
            patternFillType: 1,
            patternTileScale: 1
          }
        ],
        miterLimit: 10,
        startMarkerType: 0,
        endMarkerType: 0,
        windingRule: 1
      },
      layers: [
        {
          _class: 'rectangle',
          do_objectID: this._id(),
          booleanOperation: -1,
          edited: false,
          isClosed: true,
          pointRadiusBehaviour: 1,
          points: this._rectanglePoints(rect.width, rect.height),
          frame: this._frame(0, 0, rect.width, rect.height),
          style: this._defaultStyle()
        }
      ]
    };
  }

  _textLayer(textNode, rootRect, win) {
    const range = textNode.ownerDocument.createRange();
    range.selectNodeContents(textNode);
    const rect = range.getBoundingClientRect();
    if (!rect.width || !rect.height) {
      return null;
    }

    const parent = textNode.parentElement;
    const style = win.getComputedStyle(parent);
    const content = textNode.textContent.trim();
    if (!content) {
      return null;
    }

    const color = this._parseColor(style.color) || this._parseColor('rgba(15,23,42,1)');
    const fontSize = parseFloat(style.fontSize) || 14;
    const fontFamily = style.fontFamily.split(',')[0].replace(/['"]/g, '').trim() || 'Helvetica';

    return {
      _class: 'text',
      do_objectID: this._id(),
      name: content.slice(0, 32),
      booleanOperation: -1,
      clippingMaskMode: 0,
      exportOptions: this._emptyExportOptions(),
      frame: this._frame(
        rect.left - rootRect.left,
        rect.top - rootRect.top,
        rect.width,
        rect.height
      ),
      isLocked: false,
      isVisible: true,
      layerListExpandedType: 0,
      lineSpacingBehaviour: 2,
      nameIsFixed: false,
      resizingConstraint: 63,
      resizingType: 0,
      rotation: 0,
      shouldBreakMaskChain: false,
      style: {
        _class: 'style',
        textStyle: {
          _class: 'textStyle',
          encodedAttributes: {
            MSAttributedStringColorAttribute: color,
            MSAttributedStringFontAttribute: {
              _class: 'fontDescriptor',
              attributes: {
                name: fontFamily,
                size: fontSize
              }
            },
            paragraphStyle: {
              _class: 'paragraphStyle',
              alignment: this._textAlignment(style.textAlign)
            }
          }
        }
      },
      attributedString: {
        _class: 'attributedString',
        string: content,
        attributes: [
          {
            _class: 'stringAttribute',
            location: 0,
            length: content.length,
            attributes: {
              MSAttributedStringColorAttribute: color,
              MSAttributedStringFontAttribute: {
                _class: 'fontDescriptor',
                attributes: {
                  name: fontFamily,
                  size: fontSize
                }
              },
              paragraphStyle: {
                _class: 'paragraphStyle',
                alignment: this._textAlignment(style.textAlign)
              }
            }
          }
        ]
      },
      textBehaviour: 0
    };
  }

  _rectanglePoints(width, height) {
    return [
      {
        _class: 'curvePoint',
        cornerRadius: 0,
        curveFrom: '{0, 0}',
        curveMode: 1,
        curveTo: '{0, 0}',
        hasCurveFrom: false,
        hasCurveTo: false,
        point: '{0, 0}'
      },
      {
        _class: 'curvePoint',
        cornerRadius: 0,
        curveFrom: '{1, 0}',
        curveMode: 1,
        curveTo: '{1, 0}',
        hasCurveFrom: false,
        hasCurveTo: false,
        point: '{1, 0}'
      },
      {
        _class: 'curvePoint',
        cornerRadius: 0,
        curveFrom: '{1, 1}',
        curveMode: 1,
        curveTo: '{1, 1}',
        hasCurveFrom: false,
        hasCurveTo: false,
        point: '{1, 1}'
      },
      {
        _class: 'curvePoint',
        cornerRadius: 0,
        curveFrom: '{0, 1}',
        curveMode: 1,
        curveTo: '{0, 1}',
        hasCurveFrom: false,
        hasCurveTo: false,
        point: '{0, 1}'
      }
    ];
  }

  _frame(x, y, width, height) {
    return {
      _class: 'rect',
      constrainProportions: false,
      height: Number(height.toFixed(2)),
      width: Number(width.toFixed(2)),
      x: Number(x.toFixed(2)),
      y: Number(y.toFixed(2))
    };
  }

  _defaultStyle() {
    return {
      _class: 'style',
      miterLimit: 10,
      startMarkerType: 0,
      endMarkerType: 0,
      windingRule: 1
    };
  }

  _emptyExportOptions() {
    return {
      _class: 'exportOptions',
      exportFormats: [],
      includedLayerIds: [],
      layerOptions: 0,
      shouldTrim: false
    };
  }

  _rulerData() {
    return {
      _class: 'rulerData',
      base: 0,
      guides: []
    };
  }

  _parseColor(color) {
    if (!color) return null;
    const ctx = color.trim();
    const rgbaMatch = ctx.match(/rgba?\(([^)]+)\)/i);
    if (!rgbaMatch) {
      return {
        _class: 'color',
        alpha: 1,
        red: 1,
        green: 1,
        blue: 1
      };
    }

    const parts = rgbaMatch[1]
      .split(',')
      .map((part) => part.trim())
      .map((part, index) => (index < 3 ? parseFloat(part) / 255 : parseFloat(part)));

    const [r = 1, g = 1, b = 1, a = 1] = parts;
    return {
      _class: 'color',
      alpha: isNaN(a) ? 1 : a,
      red: isNaN(r) ? 0 : r,
      green: isNaN(g) ? 0 : g,
      blue: isNaN(b) ? 0 : b
    };
  }

  _textAlignment(alignment) {
    switch (alignment) {
      case 'right':
        return 1;
      case 'center':
        return 2;
      case 'justify':
        return 3;
      case 'left':
      default:
        return 0;
    }
  }

  _id() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID().toUpperCase();
    }

    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === 'x' ? r : (r & 0x3) | 0x8;
      const hex = v.toString(16);
      return c === 'x' ? hex : hex;
    }).toUpperCase();
  }
}
