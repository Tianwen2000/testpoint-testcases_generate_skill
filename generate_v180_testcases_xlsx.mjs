import fs from "node:fs/promises";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PYTHON = "/Users/a123/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const OUTPUT_DIR = path.join(ROOT, "outputs", "v1.8.0_testcases");
const OUTPUT_FILE = path.join(OUTPUT_DIR, "音创AI_MelodAI_V1.8.0_首页_创作页_我的_测试用例.xlsx");

const PRODUCT = "音创AI/Melod AI";
const VERSION = "V1.8.0";
const PROJECT = "音创AI/Melod AI V1.8.0";
const MODULE_PREFIX = { "首页": "SY", "创作页": "CZ", "我的": "WD" };

function loadTrees() {
  const code = `
import importlib.util, json, pathlib
root = pathlib.Path(${JSON.stringify(ROOT)})
spec = importlib.util.spec_from_file_location("gen", root / "scripts/generate_v180_testpoint_xmind.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
print(json.dumps(mod.TREES, ensure_ascii=False))
`;
  return JSON.parse(execFileSync(PYTHON, ["-c", code], { cwd: ROOT, encoding: "utf8" }));
}

function walkLeaves(item, pathParts, out) {
  const next = [...pathParts, item.title];
  if (!item.children || item.children.length === 0) {
    out.push({ title: item.title, path: next });
    return;
  }
  for (const child of item.children) walkLeaves(child, next, out);
}

function sectionOf(leaf) {
  return leaf.path[2] || leaf.path[1] || "";
}

function moduleOf(leaf) {
  return leaf.path[1] || "";
}

function branchOf(leaf) {
  const branch = leaf.path.slice(2, -1);
  return branch.length ? branch.join(" -> ") : sectionOf(leaf);
}

function pathText(leaf) {
  return leaf.path.join(" > ");
}

function isPendingTitle(title) {
  return /待确认|需确认|未定义|冲突|设计图未定义|规则待确认|处理待确认|兜底待确认|范围待确认/.test(title);
}

function nodeType(title) {
  if (isPendingTitle(title)) return "待确认";
  if (/失败|拒绝|不足|异常|弱网|重复|取消|超时|不可用|杀进程|错误/.test(title)) return "异常";
  if (/成功|回流|返回|恢复|跳转|刷新|清空|保留|退回|回显|结果/.test(title)) return "结果";
  if (/状态|排队|生成中|下载中|处理中|置灰|禁用|收起|展开/.test(title)) return "状态";
  if (/展示|显示|隐藏|默认|文案|入口|页面|列表|横幅|弹窗|气泡|Toast/.test(title)) return "展示";
  if (/点击|上传|创建|删除|下载|分享|切换|选择|输入|勾选|关闭|打开|拉起|跳转/.test(title)) return "交互";
  if (/会员|非会员|钻石|扣|消耗|权限|限制|字数|范围|配置|同步|不重复|二次确认/.test(title)) return "规则";
  return "规则";
}

function scenarioTag(title) {
  if (/失败/.test(title)) return "失败";
  if (/不足|无钻石|余额不足/.test(title)) return "资源不足";
  if (/会员/.test(title) && /非会员/.test(title)) return "会员与非会员分流";
  if (/会员/.test(title)) return "会员路径";
  if (/非会员/.test(title)) return "非会员路径";
  if (/权限/.test(title)) return "权限";
  if (/边界|2 分钟|80|200|4500|0-10|50|20 个|1.5 秒|3 秒/.test(title)) return "边界";
  if (/成功/.test(title)) return "成功";
  if (/默认/.test(title)) return "默认状态";
  if (/回流|返回|跳转|恢复/.test(title)) return "回流";
  if (/创建|上传|下载|分轨|创作|添加|删除|分享|切换|点击/.test(title)) return "交互";
  return "规则";
}

function preconditions(module, section, title) {
  const items = ["测试环境已安装 V1.8.0 版本应用"];
  if (/非会员/.test(title)) items.push("已登录非会员账号");
  else if (/会员/.test(title)) items.push("已登录会员账号");
  else items.push("已登录普通测试账号");
  if (/钻石充足|余额充足|有钻石|足够/.test(title)) items.push("账号钻石余额满足本场景消耗");
  if (/钻石不足|无钻石|余额不足/.test(title)) items.push("账号钻石余额不足或为 0");
  if (/上传|音频|分轨/.test(title)) items.push("已准备可用于测试的音频文件");
  if (/超过 2 分钟|>2|2 分钟 1 秒/.test(title)) items.push("已准备时长大于 2 分钟的音频");
  if (/不超过 2 分钟|<=2|前 2 分钟|2 分钟整点/.test(title)) items.push("已准备时长不超过 2 分钟的音频");
  if (module === "我的" || /播放页|歌曲|专辑|作品/.test(title)) items.push("账号下已存在可播放作品");
  if (/空状态|暂无可添加/.test(title)) items.push("账号处于对应空数据状态");
  if (/下载/.test(title)) items.push("当前作品支持下载入口展示");
  if (/后台配置|配置/.test(title)) items.push("后台已下发本场景所需配置");
  return [...new Set(items)].join("\n");
}

function actionSteps(module, section, title) {
  const steps = ["打开应用并登录满足前置条件的账号"];
  if (module === "首页") {
    if (section === "入口与默认落点") {
      steps.push("冷启动应用，并按标题场景执行首次进入、后台恢复或重新打开操作");
      steps.push("观察首屏落点、旧入口回归表现和纯音乐默认状态");
    } else if (section === "首页界面翻新") {
      steps.push("进入首页新版界面，滚动查看页面入口、卡片、按钮和图文区域");
      steps.push("按标题场景点击关键入口或切换屏幕尺寸进行检查");
    } else if (/订阅页|价格|付费/.test(title)) {
      steps.push("从首页触发需要订阅或付费的入口");
      steps.push("观察订阅页、价格展示和返回链路");
    } else if (section === "有钻石无会员能力边界") {
      steps.push("使用非会员账号进入首页相关付费功能入口");
      steps.push("按标题场景分别校验有钻石、无钻石和下载格式限制");
    } else if (section === "商业化与扣费") {
      steps.push("进入首页分轨链路并上传满足场景要求的音频");
      steps.push("按标题场景触发分轨、订阅或钻石消耗流程");
    } else if (section === "异常与兼容") {
      steps.push("进入首页分轨链路并按标题场景模拟权限、弱网、重复点击或配置异常");
      steps.push("观察页面拦截、提示、任务创建和上下文保留情况");
    } else {
      steps.push("在首页点击【分轨】入口");
      steps.push("按当前测试标题执行上传、分流、付费或结果查看操作");
    }
  } else if (module === "创作页") {
    steps.push("进入创作页并切换到对应的灵感创作或高级创作模式");
    if (/音频参考|上传|录制|版权|Suno/.test(title)) {
      steps.push("点击【+音频】并按场景选择本地上传、现场录制或播放页带入音频");
    } else if (/风格|标签|优化/.test(title)) {
      steps.push("在风格描述区域输入或选择标签，并触发相关按钮");
    } else if (/模型|会员|订阅/.test(title)) {
      steps.push("打开模型选择区域并选择对应模型");
    } else if (/提交|任务|成功|失败|查看|清空/.test(title)) {
      steps.push("填写满足提交条件的创作内容并点击生成");
    } else {
      steps.push("按当前测试标题填写、切换或调整对应字段");
    }
  } else {
    if (/播放页/.test(section) || /下载|分享|分轨|创作同款/.test(title)) {
      steps.push("进入作品播放页，打开对应功能入口");
      steps.push("按当前测试标题执行二创、下载、分轨、分享或删除操作");
    } else {
      steps.push("进入【我的】-【专辑】页面");
      steps.push("按当前测试标题执行创建、添加、编辑、删除或播放操作");
    }
  }
  steps.push("观察页面状态、提示文案、列表数据、余额/权益变化和跳转结果");
  return steps.map((step, index) => `${index + 1}. ${step}`).join("\n");
}

function expectedResult(module, section, title) {
  if (/字数限制/.test(title)) {
    return `${title}；达到限制值时输入仍可保存，超过限制后继续输入被拦截或不再写入，原有内容不丢失。`;
  }
  if (/范围 0-10|显示范围为 0-10/.test(title)) {
    return `${title}；低于最小值、高于最大值和超过 1 位小数的输入被限制，默认值展示正确。`;
  }
  if (/成功/.test(title)) {
    return `${title}；成功提示、目标页面、列表数据和相关状态刷新正确，不产生重复记录或脏数据。`;
  }
  if (/失败/.test(title)) {
    return `${title}；失败提示文案准确，页面停留和可重试状态正确，已定义的余额或数据回退正确。`;
  }
  if (/不足|无钻石|余额不足/.test(title)) {
    return `${title}；系统按 PRD 展示订阅或购买钻石入口，不直接执行需权益或需余额的操作。`;
  }
  if (/会员|非会员|钻石/.test(title)) {
    return `${title}；账号身份、余额判断、按钮展示、跳转入口和扣减结果均与 PRD 分流规则一致。`;
  }
  if (/下载/.test(title)) {
    return `${title}；下载任务状态、横幅提示、文件或相册跳转、失败退钻表现与 PRD 一致。`;
  }
  if (/上传|录制|权限/.test(title)) {
    return `${title}；上传入口、权限触发、上传结果和后续页面状态正确，无重复上传或上下文丢失。`;
  }
  if (/跳转|返回|回流|恢复|查看/.test(title)) {
    return `${title}；目标页面正确，来源参数和任务结果回流正确，返回后不丢失当前上下文。`;
  }
  if (/展示|显示|隐藏|默认|文案|入口|页面|列表|横幅|弹窗|气泡|Toast/.test(title)) {
    return `${title}；对应页面元素、文案、默认值、显示/隐藏状态与 PRD 要求一致。`;
  }
  if (/删除|二次确认/.test(title)) {
    return `${title}；二次确认、取消保留、确认删除和关联数据保留规则均正确。`;
  }
  return `${title}；页面状态、数据变化、按钮可用性和提示结果均与 PRD 对应规则一致。`;
}

function sourceRemark(leaf) {
  return `XMind路径：${pathText(leaf)}\n来源：音创AI/Melod AI-- V1.8.0.md；需求分析与测试理解简报_V1.8.0.md`;
}

function isAbnormal(title) {
  return /失败|拒绝|不足|异常|弱网|重复|取消|超时|不可用|杀进程|错误|无钻石/.test(title);
}

function isRecovery(title) {
  return /回流|返回|恢复|重进|后台|跳转|刷新|保留|清空|查看|自动执行|继续/.test(title);
}

function isBoundary(title) {
  return /边界|2 分钟|2分钟|80|200|4500|0-10|50|20 个|1.5 秒|3 秒|超过|不超过|小数|字数|上限|长/.test(title);
}

function isBranch(title) {
  return /会员|非会员|钻石|歌曲|纯音乐|翻唱|添加人声|添加纯音乐|MP3|MP4|WAV|现有|新建/.test(title);
}

function isIntercept(title) {
  return /触发订阅|拦截|不显示|不可|置灰|二次确认|权限|拒绝|限制|关闭|取消/.test(title);
}

function buildRows(trees) {
  const leaves = [];
  for (const tree of Object.values(trees)) walkLeaves(tree, [], leaves);

  const cases = [];
  const mappings = [];
  const pending = [];
  const caseByPath = new Map();
  const pendingByPath = new Map();
  let caseSeq = 1;
  let pendingSeq = 1;

  for (const leaf of leaves) {
    const module = moduleOf(leaf);
    const section = sectionOf(leaf);
    const pendingNode = isPendingTitle(leaf.title) || section === "待确认事项";
    if (pendingNode) {
      const pendingId = `PEND-${String(pendingSeq++).padStart(3, "0")}`;
      pendingByPath.set(pathText(leaf), pendingId);
      pending.push([
        pendingId,
        module,
        branchOf(leaf),
        leaf.title,
        "原始需求未定义、存在冲突或设计未落地",
        "澄清后再转化为正式测试用例，避免污染正式预期结果",
        sourceRemark(leaf),
      ]);
      mappings.push([
        pathText(leaf),
        nodeType(leaf.title),
        "否",
        "",
        `原始需求未定义或存在冲突，已登记为 ${pendingId}`,
      ]);
      continue;
    }

    const caseId = `TC-V180-${String(caseSeq++).padStart(3, "0")}`;
    caseByPath.set(pathText(leaf), caseId);
    const link = `${module}-${branchOf(leaf)}`;
    const title = `${section}（子模块）之${scenarioTag(leaf.title)}（场景）的验证：${leaf.title}`;
    cases.push([
      caseId,
      PRODUCT,
      link,
      VERSION,
      title,
      actionSteps(module, section, leaf.title),
      preconditions(module, section, leaf.title),
      expectedResult(module, section, leaf.title),
      "",
      sourceRemark(leaf),
      "",
      "",
    ]);
    mappings.push([
      pathText(leaf),
      nodeType(leaf.title),
      "是",
      caseId,
      "",
    ]);
  }

  return { leaves, cases, mappings, pending, caseByPath, pendingByPath };
}

function buildCoverage(trees, cases, leaves, caseByPath, pendingByPath) {
  const rows = [];
  for (const tree of Object.values(trees)) {
    const moduleNode = tree.children[0];
    for (const section of moduleNode.children) {
      const sectionLeaves = [];
      walkLeaves(section, [tree.title, moduleNode.title], sectionLeaves);
      const ids = sectionLeaves
        .map((leaf) => caseByPath.get(pathText(leaf)))
        .filter(Boolean);
      const pendingIds = sectionLeaves
        .map((leaf) => pendingByPath.get(pathText(leaf)))
        .filter(Boolean);
      const by = (fn) => sectionLeaves
        .filter((leaf) => fn(leaf.title))
        .map((leaf) => caseByPath.get(pathText(leaf)))
        .filter(Boolean);
      const main = ids.filter((id) => {
        const row = cases.find((c) => c[0] === id);
        return row && !isAbnormal(row[4]);
      });
      const abnormal = by(isAbnormal);
      const recovery = by(isRecovery);
      const boundary = by(isBoundary);
      const branch = by(isBranch);
      const intercept = by(isIntercept);

      const valueOrReason = (arr, reason) => arr.length ? arr.join("、") : reason;
      const pendingReason = pendingIds.length ? `；相关待确认：${pendingIds.join("、")}` : "";
      rows.push([
        section.title,
        `${moduleNode.title}-${section.title}`,
        valueOrReason(main, `不适用：该核心点在 XMind 中未拆出独立主流程叶子节点${pendingReason}`),
        valueOrReason(abnormal, `不适用：PRD 未定义该核心点的关键异常流程${pendingReason}`),
        valueOrReason(recovery, `不适用：该核心点无跨页回流、恢复或重进要求${pendingReason}`),
        valueOrReason(boundary, `不适用：该核心点无数值、时长、数量或文本边界要求${pendingReason}`),
        valueOrReason(branch, `不适用：该核心点无会员、资源、模式、格式或入口正反分支${pendingReason}`),
        valueOrReason(intercept, `不适用：该核心点无明确拦截条件${pendingReason}`),
        ids.length ? ids.join("、") : "无正式用例",
        pendingIds.length ? `存在待确认事项：${pendingIds.join("、")}` : "无",
      ]);
    }
  }
  return rows;
}

function buildStats(trees, cases, mappings, pending) {
  const rows = [["项目", "数值", "说明"]];
  rows.push(["所属产品", PRODUCT, PROJECT]);
  rows.push(["测试版本", VERSION, "基于 V1.8.0 PRD、需求分析简报和 XMind 测试点"]);
  rows.push(["正式测试用例数", cases.length, "不含待确认事项"]);
  rows.push(["XMind 叶子节点数", mappings.length, "每个叶子节点均有映射"]);
  rows.push(["待确认事项数", pending.length, "未写入正式预期"]);
  for (const module of Object.keys(trees)) {
    const moduleCases = cases.filter((row) => row[2].startsWith(`${module}-`)).length;
    const moduleMappings = mappings.filter((row) => row[0].includes(`> ${module} >`)).length;
    const modulePending = pending.filter((row) => row[1] === module).length;
    rows.push([`${module} 用例数`, moduleCases, `叶子节点 ${moduleMappings} 个，待确认 ${modulePending} 个`]);
  }
  return rows;
}

function colName(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function writeSheet(workbook, name, headers, rows, widths) {
  const sheet = workbook.worksheets.add(name);
  const data = [headers, ...rows];
  const lastCol = colName(headers.length);
  const range = sheet.getRange(`A1:${lastCol}${data.length}`);
  range.values = data;
  range.format.wrapText = true;
  range.format.verticalAlignment = "top";
  range.format.font = { name: "Arial", size: 10 };
  range.format.borders = { preset: "all", style: "thin", color: "#D9E2EC" };
  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format.fill = "#1F4E78";
  header.format.font = { color: "#FFFFFF", bold: true, size: 10 };
  header.format.horizontalAlignment = "center";
  header.format.rowHeightPx = 34;
  for (let i = 0; i < widths.length; i++) {
    sheet.getRange(`${colName(i + 1)}:${colName(i + 1)}`).format.columnWidthPx = widths[i];
  }
  sheet.getRange(`A2:${lastCol}${data.length}`).format.rowHeightPx = 96;
  return sheet;
}

async function main() {
  const trees = loadTrees();
  const workbook = Workbook.create();
  const built = buildRows(trees);
  const coverage = buildCoverage(trees, built.cases, built.leaves, built.caseByPath, built.pendingByPath);
  const stats = buildStats(trees, built.cases, built.mappings, built.pending);

  writeSheet(
    workbook,
    "测试用例",
    ["用例编号", "所属产品", "所属模块/链路", "测试版本", "测试标题", "操作步骤", "前置条件", "预期结果", "实际结果", "附件（备注）", "测试负责人", "时间"],
    built.cases,
    [100, 130, 220, 80, 380, 420, 260, 460, 100, 360, 100, 100],
  );
  writeSheet(
    workbook,
    "核心测试点覆盖矩阵",
    ["核心测试点", "所属模块/链路", "主流程", "关键异常", "关键回流/恢复", "关键边界", "关键正反分支", "关键拦截条件", "对应测试用例编号", "不适用说明"],
    coverage,
    [180, 220, 360, 360, 360, 360, 360, 360, 420, 300],
  );
  writeSheet(
    workbook,
    "XMind叶子节点映射",
    ["XMind 原节点路径", "节点类型（展示/交互/规则/状态/异常/结果/全局）", "是否生成独立用例", "对应测试用例编号", "未独立生成原因"],
    built.mappings,
    [520, 180, 120, 180, 420],
  );
  writeSheet(
    workbook,
    "待确认事项清单",
    ["待确认编号", "所属模块", "所属链路", "待确认事项", "原因类型", "处理建议", "来源备注"],
    built.pending,
    [110, 100, 220, 460, 220, 420, 420],
  );
  writeSheet(
    workbook,
    "统计",
    ["项目", "数值", "说明"],
    stats.slice(1),
    [220, 160, 520],
  );

  const inspect = await workbook.inspect({
    kind: "table",
    range: "测试用例!A1:H8",
    include: "values",
    tableMaxRows: 8,
    tableMaxCols: 8,
  });
  console.log(inspect.ndjson.split("\n").slice(0, 3).join("\n"));
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "formula error scan",
  });
  console.log(errors.ndjson || "formula_error_scan_empty");
  for (const sheetName of ["测试用例", "核心测试点覆盖矩阵", "XMind叶子节点映射", "待确认事项清单", "统计"]) {
    await workbook.render({ sheetName, range: "A1:E12", scale: 1 });
    console.log(`render_ok:${sheetName}`);
  }

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(OUTPUT_FILE);
  console.log(JSON.stringify({
    output: OUTPUT_FILE,
    cases: built.cases.length,
    mappings: built.mappings.length,
    pending: built.pending.length,
    coverage: coverage.length,
  }, null, 2));
}

main();
