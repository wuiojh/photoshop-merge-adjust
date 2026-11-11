#target photoshop
// merge_adjust_v20 —— 顶部调色层逐个合并进下方主体图层；图层原名 + 剪贴蒙版结构原样保留
//
// 【v20 关键修正：只合并「纯像素图层」，其余一律跳过】
//  v18/v19 曾把 SOLIDFILL / GRADIENTFILL / PATTERNFILL（形状、填充图层）也当成主体层，
//  导致带矢量蒙版/智能特征、名字形如「形状 1」的图层被合并栅格化 —— 这就是它没被跳过的原因。
//  现在改成【白名单】：只有 LayerKind.NORMAL 且 Action Manager 复核为纯像素的图层才参与合并。
//  智能对象、形状图层、填充图层、文字图层、调整图层、3D/视频图层全部跳过，
//  跳过时记录名字+类型，收尾按 ID 复核它们仍然存在、名字未变、未被栅格化。
//
// 【v19 新增：智能对象（智能图层）硬性跳过】
//  遇到智能对象一律跳过：不合并、不栅格化、不改名、不进任何兜底流程。
//  双重识别：先按 LayerKind.SMARTOBJECT 判断；某些版本智能对象的 kind 会回落成 NORMAL，
//            再用 Action Manager 读图层描述符里的 smartObject 键复核，结果按 ID 缓存。
//  收尾还会按 ID 复核这些智能对象仍然存在、名字未变，确保它们"完全没被碰过"。
//
// 【v18 新增：剪贴蒙版（剪贴组）保护】
//  情形一：色相/饱和度、曲线等【调整图层】剪贴到主体图层上（如「色相/饱和度 1」→「形状 2」）
//          → 调整图层完全跳过：不合并、不改名、不动它的蒙版；
//            先释放剪贴关系，只把调色层合并进下面的主体图层，
//            合并完成后把该调整图层重新剪贴回「原本的主体图层」，复原原样。
//  情形二：上方是【普通像素剪贴图层】（如「图层 15」剪贴到「图层 9」）
//          → 先释放剪贴关系，让剪贴图层与主体图层各自独立地"分别"与调色层合并，
//            合并完成后把原来的剪贴蒙版重新创建回去，还原原来的上下结构。
//  实现：开场按【图层 ID】快照 名字 + grouped 状态 → 全部释放 → 分别合并
//        → 按 ID / 结果层映射精确重建 grouped=true（自下而上），并逐个回读校验。
//        即使合并结果换成了新的图层对象，也会继承原目标层该有的剪贴状态。
//
// 【v17 已修复：运行完所有图层名字被改掉】按 ID 快照名字、只按 ID 精确改名、
//  不再按位置批量还原、不再解锁背景层、不再给复制层设 grouped=true 制造默认名、
//  剪贴板方案用可见性快照精确恢复、单个名字异常不再中止全脚本。
//
// 规则：✔ 只有可见的【纯像素图层】→ 合并调色层；✘ 图层组只钻进去找纯像素层；
//       ✘ 智能对象 / 形状图层 / 填充图层 / 调整图层 / 文字 / 隐藏层 → 全部跳过并保持原样。
// ⚠ 合并会栅格化形状与像素内容，请先在 PSD 副本上运行！

if (app.documents.length === 0) {
    alert("请先打开一个文档");
} else {
    var doc = app.activeDocument;
    var adj = doc.activeLayer;
    var oldDialogs = app.displayDialogs;

    function focusDoc() { try { app.activeDocument = doc; } catch (e) {} }

    function killTemps() {
        for (var d = app.documents.length - 1; d >= 0; d--) {
            var nm = "";
            try { nm = String(app.documents[d].name); } catch (e) { continue; }
            if (nm.indexOf("TMP_MERGE") === 0) {
                try {
                    var sv = app.displayDialogs;
                    app.displayDialogs = DialogModes.NO;
                    app.documents[d].close(SaveOptions.DONOTSAVE);
                    app.displayDialogs = sv;
                } catch (e) {}
            }
        }
    }

    // 按 ID 精确改名：绝不使用 "Trgt"（当前活动图层），避免误伤其它图层
    function forceName(lay, nm) {
        if (!nm) return false;
        try { if (lay.name === nm) return true; } catch (e) { return false; }
        try {
            lay.name = nm;
            if (lay.name === nm) return true;
        } catch (e) {}
        try {
            var r = new ActionReference();
            r.putIdentifier(stringIDToTypeID("layer"), lay.id);
            var d = new ActionDescriptor();
            d.putReference(charIDToTypeID("null"), r);
            var p = new ActionDescriptor();
            p.putString(charIDToTypeID("Nm  "), String(nm));
            d.putObject(charIDToTypeID("T   "), charIDToTypeID("Lyr "), p);
            executeAction(charIDToTypeID("setd"), d, DialogModes.NO);
        } catch (e) {}
        try { return (lay.name === nm); } catch (e) { return false; }
    }

    // 调整图层（色相/饱和度、曲线、色彩平衡…）：必须原样跳过
    function isAdjustmentLayer(L) {
        try {
            var k = String(L.kind);
            var kinds = ["LEVELS", "CURVES", "COLORBALANCE", "BRIGHTNESSCONTRAST",
                         "HUESATURATION", "SELECTIVECOLOR", "CHANNELMIXER", "GRADIENTMAP",
                         "INVERT", "THRESHOLD", "POSTERIZE", "PHOTOFILTER", "SHADOWHIGHLIGHT",
                         "EXPOSURE", "VIBRANCE", "BLACKANDWHITE", "COLORLOOKUP"];
            for (var i = 0; i < kinds.length; i++) if (k.indexOf(kinds[i]) >= 0) return true;
        } catch (e) {}
        return false;
    }

    // 读取图层的 Action Manager 描述符（按 ID），失败返回 null
    function layerDesc(L) {
        try {
            var r = new ActionReference();
            r.putIdentifier(charIDToTypeID("Lyr "), L.id);
            return executeActionGet(r);
        } catch (e) { return null; }
    }

    // 图层类型标签（用于跳过明细，便于人工核对）
    function describeKind(L) {
        try {
            var k = String(L.kind).toUpperCase();
            if (k.indexOf("SMART") >= 0) return "智能对象";
            if (k.indexOf("SOLIDFILL") >= 0 || k.indexOf("GRADIENTFILL") >= 0 ||
                k.indexOf("PATTERNFILL") >= 0) return "形状/填充图层";
            if (k.indexOf("TEXT") >= 0) return "文字图层";
            if (isAdjustmentLayer(L)) return "调整图层";
            if (k.indexOf("NORMAL") >= 0) return "像素图层";
            if (k.indexOf("LAYER3D") >= 0 || k.indexOf("VIDEO") >= 0) return "3D/视频图层";
            return "特殊图层";
        } catch (e) { return "特殊图层"; }
    }

    // 智能对象识别：多重兜底，绝不漏判
    //  ① kind 含 SMART       ② 描述符含 smartObject / smartObjectMore
    //  ③ layerKind === 5     ④ 显式 putProperty 请求 smartObject
    var smartCache = {};
    function detectSmart(L) {
        try {
            if (L.typename === "LayerSet") return false;
            if (String(L.kind).toUpperCase().indexOf("SMART") >= 0) return true;
            var d = layerDesc(L);
            if (d !== null) {
                if (d.hasKey(stringIDToTypeID("smartObject")) ||
                    d.hasKey(stringIDToTypeID("smartObjectMore"))) return true;
                if (d.hasKey(stringIDToTypeID("layerKind")) &&
                    d.getInteger(stringIDToTypeID("layerKind")) === 5) return true;
            }
            try {
                var r2 = new ActionReference();
                r2.putProperty(charIDToTypeID("Prpr"), stringIDToTypeID("smartObject"));
                r2.putIdentifier(charIDToTypeID("Lyr "), L.id);
                if (executeActionGet(r2).hasKey(stringIDToTypeID("smartObject"))) return true;
            } catch (e) {}
        } catch (e) {}
        return false;
    }
    function isSmartObject(L) {
        var id = "";
        try { id = String(L.id); } catch (e) {}
        if (smartCache[id] !== undefined) return smartCache[id];
        var yes = detectSmart(L);
        smartCache[id] = yes;
        return yes;
    }

    // 【白名单】只有纯像素图层可以合并
    //  硬门槛：kind 必须是 NORMAL（形状/填充/文字/智能/调整层在这里就被挡掉）
    //  复核：AM 的 layerKind 落在已知「非像素」区间 2..11 时也排除
    //        （2 调整 / 3 文字 / 4 矢量形状 / 5 智能对象 / 6 视频 / 7 组 /
    //         8 3D / 9 渐变填充 / 10 图案填充 / 11 纯色填充）
    function isPixelLayer(L) {
        try {
            if (L.typename === "LayerSet") return false;
            if (isAdjustmentLayer(L)) return false;
            if (isSmartObject(L)) return false;
            if (String(L.kind).toUpperCase().indexOf("NORMAL") < 0) return false;
            var d = layerDesc(L);
            if (d !== null && d.hasKey(stringIDToTypeID("layerKind"))) {
                var lk = d.getInteger(stringIDToTypeID("layerKind"));
                if (lk >= 2 && lk <= 11) return false;
            }
            return true;
        } catch (e) {}
        return false;
    }

    // 记录目标在父容器中的【位置索引】；回拷时重新读取容器，避免引用失效
    function getAnchor(t) {
        try {
            var p = t.parent;
            if (p === null || p === undefined) p = doc;
            var kids = p.layers;
            for (var i = 0; i < kids.length; i++) {
                if (kids[i] === t) return { parent: p, index: i };
            }
        } catch (e) {}
        return null;
    }
    function placeByIndex(layer, anchor) {
        if (anchor === null) return;
        try {
            var kids = anchor.parent.layers;   // 移除原图层后重新取，引用都是新的
            var i = anchor.index;
            if (i - 1 >= 0 && i - 1 < kids.length) {
                layer.move(kids[i - 1], ElementPlacement.PLACEAFTER);
            } else if (i < kids.length) {
                layer.move(kids[i], ElementPlacement.PLACEBEFORE);
            }
        } catch (e) {}
    }

    function snapshotVisible() {
        var rec = [];
        function walk(c) {
            for (var i = 0; i < c.layers.length; i++) {
                var L = c.layers[i];
                try { rec.push({ layer: L, visible: L.visible }); } catch (e) {}
                if (L.typename === "LayerSet") walk(L);
            }
        }
        walk(doc);
        return rec;
    }
    function restoreVisible(rec) {
        for (var i = 0; i < rec.length; i++) {
            try { rec[i].layer.visible = rec[i].visible; } catch (e) {}
        }
    }
    function hideAll(container) {
        for (var i = 0; i < container.layers.length; i++) {
            var L = container.layers[i];
            try { L.visible = false; } catch (e) {}
            if (L.typename === "LayerSet") hideAll(L);
        }
    }
    function showWithParents(L) {
        var p = L;
        for (var g = 0; p && p !== doc && g < 50; g++) {
            try { p.visible = true; } catch (e) {}
            try { p = p.parent; } catch (e) { break; }
        }
    }
    function restoreLook(dst, op, bm) {
        try { dst.opacity = op; } catch (e) {}
        try { dst.blendMode = bm; } catch (e) {}
    }

    // ===== 剪贴蒙版：开场全部释放，收尾原样重建 =====
    function releaseClips(container) {
        var n = 0;
        for (var i = container.layers.length - 1; i >= 0; i--) {
            var L = container.layers[i];
            if (L.typename === "LayerSet") n += releaseClips(L);
            try { if (L.grouped) { L.grouped = false; n++; } } catch (e) {}
        }
        return n;
    }
    var badClips = [];
    function restoreClips(container, wantGrouped) {
        var n = 0;
        // 自下而上：基底先就位，上方剪贴层才能正确挂回去
        for (var i = container.layers.length - 1; i >= 0; i--) {
            var L = container.layers[i];
            var want = wantGrouped(L);
            try {
                if (want && !L.grouped) L.grouped = true;
                else if (!want && L.grouped) L.grouped = false;
            } catch (e) {}
            if (want) {
                var ok = false;
                try { ok = (L.grouped === true); } catch (e) {}
                if (ok) n++;
                else { try { badClips.push(L.name); } catch (e) { badClips.push("(未知图层)"); } }
            }
            if (L.typename === "LayerSet") n += restoreClips(L, wantGrouped);
        }
        return n;
    }

    if (adj.isBackgroundLayer || adj.typename === "LayerSet") {
        app.displayDialogs = oldDialogs;
        alert("请先在图层面板选中【最顶部那个调色图层】，再运行脚本。");
    } else {
        app.displayDialogs = DialogModes.NO;
        killTemps();
        var adjIsSmart = detectSmart(adj);   // 仅用于提示：源层是智能对象时其副本会被栅格化

        // ===== 开场快照：图层 ID -> {原名, 是否为剪贴蒙版} =====
        var nameById = {}, groupedById = {};
        function snapAll(container) {
            for (var i = 0; i < container.layers.length; i++) {
                var L = container.layers[i];
                try { nameById[String(L.id)] = L.name; } catch (e) {}
                try { groupedById[String(L.id)] = (L.grouped === true); }
                catch (e) { groupedById[String(L.id)] = false; }
                if (L.typename === "LayerSet") snapAll(L);
            }
        }
        snapAll(doc);

        // ===== 收集目标：非纯像素层全部登记为「受保护」，绝不合并 =====
        var targets = [], origNames = [], origOpacity = [], origBlend = [], origGrouped = [];
        var skipped = [], groupCount = 0, adjCount = 0, smartCount = 0, shapeCount = 0;
        var protectIds = [], protectNames = [], protectTypes = [];
        function protect(L, typeLabel) {
            try { protectIds.push(String(L.id)); } catch (e) { protectIds.push(""); }
            try { protectNames.push(L.name); } catch (e) { protectNames.push(""); }
            protectTypes.push(typeLabel);
            try { skipped.push(L.name + "（" + typeLabel + "·跳过）"); } catch (e) {}
        }
        function collect(container, parentVisible) {
            for (var i = 0; i < container.layers.length; i++) {
                var L = container.layers[i];
                if (L === adj) continue;
                if (L.typename === "LayerSet") {
                    if (!parentVisible || !L.visible) {
                        try { skipped.push(L.name + "（隐藏组）"); } catch (e) {}
                        continue;
                    }
                    groupCount++; collect(L, true);
                    continue;
                }
                // 1) 智能对象（含隐藏的）：硬性跳过，不进任何合并流程
                if (isSmartObject(L)) {
                    smartCount++; protect(L, "智能对象"); continue;
                }
                var typ = describeKind(L);
                // 2) 形状/填充、文字、调整、矢量等非纯像素层：一律跳过
                if (typ !== "像素图层" || !isPixelLayer(L)) {
                    if (typ === "调整图层") adjCount++;
                    else if (typ === "形状/填充图层") shapeCount++;
                    protect(L, (typ === "像素图层" ? "矢量/特殊图层" : typ));
                    continue;
                }
                // 3) 隐藏的纯像素层：跳过
                if (!parentVisible || !L.visible) {
                    try { skipped.push(L.name + "（隐藏）"); } catch (e) {}
                    continue;
                }
                // 4) 只有纯像素层才进目标列表
                targets.push(L);
                try { origNames.push(L.name); } catch (e) { origNames.push(""); }
                try { origOpacity.push(L.opacity); } catch (e) { origOpacity.push(100); }
                try { origBlend.push(L.blendMode); } catch (e) { origBlend.push(BlendMode.NORMAL); }
                try { origGrouped.push(groupedById[String(L.id)] === true); }
                catch (e) { origGrouped.push(false); }
            }
        }
        collect(doc, true);

        if (targets.length === 0) {
            killTemps(); focusDoc(); app.displayDialogs = oldDialogs;
            alert("没有找到可合并的纯像素图层，未做任何改动。\n智能对象 " + smartCount
                + " 个、形状/填充图层 " + shapeCount + " 个、调整图层 " + adjCount
                + " 个均按规则跳过；剪贴蒙版保持原样。\n\n跳过明细：\n"
                + skipped.slice(0, 8).join("\n"));
        } else {
            // 先释放全部剪贴蒙版：避免插入复制层 + 向下合并时切断剪贴链
            var releasedClips = releaseClips(doc);

            var resultLayers = [], resultNames = [], resultGrouped = [], report = [];
            var methodCount = { A: 0, B: 0, C: 0 };

            // 自下而上处理，避免上方图层变动影响下方引用
            for (var i = targets.length - 1; i >= 0; i--) {
                var t = targets[i];
                var name = origNames[i];
                var errA = "", errB = "", errC = "";
                var copy = null, done = false;

                try {
                    try { if (t.allLocked) t.allLocked = false; } catch (e) {}
                    try { t.visible = true; } catch (e) {}

                    // ===== A：原文档内向下合并，结果天然沿用下层名字 =====
                    try {
                        copy = adj.duplicate();
                        copy.move(t, ElementPlacement.PLACEBEFORE);
                        try { copy.grouped = false; } catch (e) {}   // 复制层不参与任何剪贴组
                        copy.visible = true;
                        t.visible = true;
                        doc.activeLayer = copy;
                        var m = copy.merge();
                        if (!forceName(m, name)) report.push("「" + name + "」合并后名字未能还原");
                        resultLayers.push(m); resultNames.push(name); resultGrouped.push(origGrouped[i]);
                        methodCount.A++; done = true;
                    } catch (eA) {
                        errA = eA.message;
                        if (copy !== null) {
                            // 图层可能已在合并中被销毁，先探测再删除
                            try { var cn = copy.name; copy.remove(); } catch (e) {}
                        }
                        copy = null;
                    }

                    // ===== B：临时文档里合并后回拷 =====
                    if (!done) {
                        var tmp = null;
                        try {
                            try {
                                tmp = app.documents.add(doc.width, doc.height, doc.resolution, "TMP_MERGE",
                                    NewDocumentMode.RGB, DocumentFill.TRANSPARENT, 1, doc.bitsPerChannel);
                            } catch (e) {
                                tmp = app.documents.add(doc.width, doc.height, doc.resolution, "TMP_MERGE",
                                    NewDocumentMode.RGB, DocumentFill.TRANSPARENT, 1, BitsPerChannelType.EIGHT);
                            }
                            var tl = t.duplicate(tmp, ElementPlacement.PLACEATEND);
                            var cl = adj.duplicate(tmp, ElementPlacement.PLACEATEND);
                            cl.move(tl, ElementPlacement.PLACEBEFORE);
                            try { cl.grouped = false; } catch (e) {}
                            cl.visible = true; tl.visible = true;
                            tmp.activeLayer = cl;
                            var bmL = cl.merge();
                            try { bmL.name = name; } catch (e) {}

                            var anc = getAnchor(t);
                            var back = bmL.duplicate(doc, ElementPlacement.PLACEATEND);
                            t.remove();
                            placeByIndex(back, anc);
                            restoreLook(back, origOpacity[i], origBlend[i]);
                            if (!forceName(back, name)) report.push("「" + name + "」回拷后名字未能还原");
                            resultLayers.push(back); resultNames.push(name); resultGrouped.push(origGrouped[i]);
                            methodCount.B++; done = true;
                        } catch (eB) {
                            errB = eB.message;
                        } finally {
                            if (tmp !== null) {
                                try {
                                    var sv2 = app.displayDialogs;
                                    app.displayDialogs = DialogModes.NO;
                                    tmp.close(SaveOptions.DONOTSAVE);
                                    app.displayDialogs = sv2;
                                } catch (e) {}
                            }
                            focusDoc();
                        }
                    }

                    // ===== C：剪贴板（合并可见），仅在 A/B 都失败时兜底 =====
                    if (!done) {
                        var vis = null;
                        try {
                            vis = snapshotVisible();
                            hideAll(doc);
                            showWithParents(t);
                            showWithParents(adj);
                            focusDoc();
                            doc.selection.selectAll();
                            doc.selection.copy(true);
                            try { doc.selection.deselect(); } catch (e) {}
                            doc.paste();
                            var pl = doc.activeLayer;
                            pl.move(t, ElementPlacement.PLACEBEFORE);
                            t.remove();
                            restoreLook(pl, origOpacity[i], origBlend[i]);
                            if (!forceName(pl, name)) report.push("「" + name + "」粘贴后名字未能还原");
                            resultLayers.push(pl); resultNames.push(name); resultGrouped.push(origGrouped[i]);
                            methodCount.C++; done = true;
                        } catch (eC) {
                            errC = eC.message;
                        } finally {
                            if (vis !== null) restoreVisible(vis);
                            focusDoc();
                        }
                    }

                    if (!done) {
                        report.push("「" + name + "」失败  A:" + (errA || "-") + " | B:" + (errB || "-") + " | C:" + (errC || "-"));
                    }
                } catch (e) {
                    report.push("「" + name + "」外层异常：" + e.message);
                }
            }

            // ===== 收尾 =====
            var total = methodCount.A + methodCount.B + methodCount.C;
            if (total > 0) { try { adj.remove(); } catch (e) {} }

            var fixed = 0, bad = [];

            // 1) 合并结果按记录的名字精确修正（新图层 ID 不在开场快照中）
            for (var r = 0; r < resultLayers.length; r++) {
                try {
                    if (resultLayers[r].name !== resultNames[r]) {
                        if (forceName(resultLayers[r], resultNames[r])) fixed++;
                        else bad.push("应为「" + resultNames[r] + "」");
                    }
                } catch (e) {}
            }

            // 2) 全量校验：只按 ID 对照开场快照，只修名字确实变了的图层
            function verify(container) {
                for (var i = 0; i < container.layers.length; i++) {
                    var L = container.layers[i];
                    try {
                        var want = nameById[String(L.id)];
                        if (want !== undefined && want !== "" && L.name !== want) {
                            if (forceName(L, want)) fixed++;
                            else bad.push("应为「" + want + "」，现为「" + L.name + "」");
                        }
                    } catch (e) {}
                    if (L.typename === "LayerSet") verify(L);
                }
            }
            verify(doc);

            // 3) 把剪贴蒙版「创建回去」：合并结果层继承原目标层的 grouped 状态
            var newGroupById = {};
            for (var rg = 0; rg < resultLayers.length; rg++) {
                try { newGroupById[String(resultLayers[rg].id)] = resultGrouped[rg]; } catch (e) {}
            }
            function wantGrouped(L) {
                var id = "";
                try { id = String(L.id); } catch (e) { return false; }
                if (newGroupById[id] !== undefined) return newGroupById[id] === true;
                return groupedById[id] === true;
            }
            var expectedClips = 0;
            function countWant(container) {
                for (var i = 0; i < container.layers.length; i++) {
                    var L = container.layers[i];
                    if (wantGrouped(L)) expectedClips++;
                    if (L.typename === "LayerSet") countWant(L);
                }
            }
            countWant(doc);
            var clipsBack = restoreClips(doc, wantGrouped);

            // 4) 复核所有「受保护图层」完好无损：仍在、名字未变、未被栅格化
            focusDoc();                 // AM 读取前确保目标文档是活动文档
            var finalById = {};
            function indexAll(container) {
                for (var i = 0; i < container.layers.length; i++) {
                    var L = container.layers[i];
                    try { finalById[String(L.id)] = L; } catch (e) {}
                    if (L.typename === "LayerSet") indexAll(L);
                }
            }
            indexAll(doc);
            var protectBad = [];
            for (var pi = 0; pi < protectIds.length; pi++) {
                var f = finalById[protectIds[pi]];
                var tag = protectNames[pi] + "（" + protectTypes[pi] + "）";
                if (f === undefined) { protectBad.push(tag + " 不见了"); continue; }
                try {
                    if (f.name !== protectNames[pi]) { protectBad.push(tag + " 名字变了"); continue; }
                    // 这些类型绝不能变成普通像素层（说明被栅格化/合并了）
                    var nowKind = describeKind(f);
                    if (protectTypes[pi] === "智能对象" && !detectSmart(f))
                        protectBad.push(tag + " 智能特征复核失败，请人工确认");
                    else if (protectTypes[pi] !== "像素图层" && nowKind === "像素图层")
                        protectBad.push(tag + " 已被栅格化");
                } catch (e) {}
            }

            killTemps(); focusDoc(); app.displayDialogs = oldDialogs;

            var msg = "完成 " + total + " / " + targets.length + " 个纯像素图层"
                    + "\n（A 直接合并 " + methodCount.A + "，B 临时文档 " + methodCount.B + "，C 剪贴板 " + methodCount.C + "）";
            if (groupCount > 0) msg += "\n已钻进 " + groupCount + " 个图层组（组结构保留）";
            if (skipped.length > 0) msg += "\n跳过 " + skipped.length + " 个图层（智能对象 / 形状填充 / 调整 / 文字 / 隐藏）";
            if (smartCount > 0) msg += "\n　智能对象 " + smartCount + " 个：跳过，未合并 / 未栅格化";
            if (shapeCount > 0) msg += "\n　形状/填充图层 " + shapeCount + " 个：跳过，未合并 / 未栅格化";
            if (adjCount > 0) msg += "\n　调整图层 " + adjCount + " 个：跳过且原样保留";
            if (releasedClips > 0 || expectedClips > 0)
                msg += "\n剪贴蒙版：释放 " + releasedClips + " 个 → 重建 " + clipsBack + " / " + expectedClips + " 个";
            msg += total > 0 ? "\n原调色层已移除" : "\n⚠ 未处理任何图层，原调色层已保留";
            if (adjIsSmart) msg += "\n⚠ 选中的调色层本身是智能对象，它的复制层在合并时会被栅格化";
            msg += "\n名字校验：修正 " + fixed + " 处";
            if (resultNames.length > 0) msg += "\n\n—— 实际合并（前 8 个）——\n" + resultNames.slice(0, 8).join("\n");
            if (skipped.length > 0) msg += "\n\n—— 跳过明细（前 8 个）——\n" + skipped.slice(0, 8).join("\n");
            if (bad.length > 0) msg += "\n\n名字未还原：\n" + bad.slice(0, 10).join("\n");
            if (protectBad.length > 0) msg += "\n\n⚠ 受保护图层异常：\n" + protectBad.slice(0, 8).join("\n");
            if (badClips.length > 0) msg += "\n\n⚠ 剪贴蒙版未能重建：\n" + badClips.slice(0, 8).join("\n");
            if (report.length > 0) msg += "\n\n详情：\n" + report.slice(0, 8).join("\n");
            alert(msg);
        }
    }
}
