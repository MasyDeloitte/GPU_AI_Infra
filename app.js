const STORAGE_KEY = 'forge-ai-sizing-projects';
const currentProjectKey = 'forge-ai-sizing-current';
const catalogRecordsKey = 'forge-ai-sizing-catalog-records';
const catalogVersion = '2026-09-16';
const catalogSource = 'AI_Sizing_infraSizing_industry_Baseline_v6.2.xlsx';
const inputIds = ['workloadType', 'queryComplexity', 'responseDetail', 'conversationLength', 'retrievalDepth', 'agenticComplexity', 'toolCalls', 'languageCoverage', 'knowledgeBase', 'concurrentSessions', 'requestsPerSecond', 'p95Ttft', 'requiredGenerationSpeed', 'precision', 'availability', 'haServers', 'drRequirement', 'drCapacity', 'planningPeriod', 'workloadGrowth', 'dataGrowth', 'serverUnitPrice', 'storageCostPerTb', 'nvmeCostPerTb', 'networkSwitchCost', 'supportRate', 'contingencyRate'];
const governed = { systemPromptTokens: 600, formattingOverheadPerTurn: 20, averageRetrievedChunkTokens: 512, averageToolResultTokens: 150, validatedModelCallsPerRequest: 2, requiredGenerationSpeed: 20 };
const models = {
  'MOD-001': { name: 'Llama 3.1 8B Instruct', context: 131072, languages: 'English, German, French, Italian, Portuguese, Hindi, Spanish, Thai' }, 'MOD-002': { name: 'Llama 3.1 70B Instruct', context: 131072, languages: 'English, German, French, Italian, Portuguese, Hindi, Spanish, Thai' }, 'MOD-003': { name: 'Llama 3.3 70B Instruct', context: 131072, languages: 'English, German, French, Italian, Portuguese, Hindi, Spanish, Thai' },
  'MOD-004': { name: 'Qwen2.5 7B Instruct', context: 32768, languages: 'Multilingual' }, 'MOD-005': { name: 'Qwen2.5 14B Instruct', context: 32768, languages: 'Multilingual' }, 'MOD-006': { name: 'Qwen2.5 32B Instruct', context: 32768, languages: 'Multilingual' },
  'MOD-008': { name: 'Mistral Small 3 24B Instruct', context: 32768, languages: 'Multilingual' }, 'MOD-009': { name: 'Mistral Small 3.1 24B Instruct', context: 131072, languages: 'Multilingual' }, 'MOD-010': { name: 'NV-Embed-e5-v5', context: 'Data Required', languages: 'Multilingual' }
};
const mappings = {
  'MOD-001|WRK-002': { id: 'MGPU-001', gpu: 'L4', vram: 24, precision: 'BF16', weight: 16, tp: 1, cmp: 'CMP-010', benchmark: 'BM-PROXY-L4' },
  'MOD-004|WRK-002': { id: 'MGPU-003', gpu: 'L40S', vram: 48, precision: 'BF16', weight: 14, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-005|WRK-003': { id: 'MGPU-004', gpu: 'L40S', vram: 48, precision: 'BF16', weight: 28, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-008|WRK-003': { id: 'MGPU-005', gpu: 'L40S', vram: 48, precision: 'INT8', weight: 24, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-006|WRK-003': { id: 'MGPU-007', gpu: 'H100 PCIe 80GB', vram: 80, precision: 'INT8', weight: 32, tp: 1, cmp: 'CMP-002', benchmark: 'Benchmark Required' },
  'MOD-002|WRK-004': { id: 'MGPU-008', gpu: 'H100 PCIe 80GB', vram: 80, precision: 'FP8', weight: 70, tp: 1, cmp: 'CMP-002', benchmark: 'Benchmark Required' },
  'MOD-003|WRK-004': { id: 'MGPU-009', gpu: 'H100 SXM5 80GB', vram: 80, precision: 'FP8', weight: 70, tp: 2, cmp: 'CMP-003', benchmark: 'BM-H100-5K500-C50' },
  'MOD-009|WRK-005': { id: 'MGPU-006', gpu: 'L40S', vram: 48, precision: 'INT8', weight: 24, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-003|WRK-005': { id: 'MGPU-010', gpu: 'H200 SXM5 141GB', vram: 141, precision: 'FP8', weight: 70, tp: 2, cmp: 'CMP-005', benchmark: 'BM-PROXY-H200' },
  'MOD-002|WRK-006': { id: 'MGPU-014', gpu: 'H200 NVL 141GB', vram: 141, precision: 'FP8', weight: 70, tp: 1, cmp: 'CMP-009', benchmark: 'BM-PROXY-H200' },
  'MOD-003|WRK-006': { id: 'MGPU-021', gpu: 'H100 SXM5 80GB', vram: 80, precision: 'FP8', weight: 70, tp: 2, cmp: 'CMP-003', benchmark: 'BM-H100-5K500-C50' },
  'MOD-008|WRK-007': { id: 'MGPU-022', gpu: 'L40S', vram: 48, precision: 'INT8', weight: 24, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-001|WRK-008': { id: 'MGPU-015', gpu: 'L40S', vram: 48, precision: 'BF16', weight: 16, tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-005|WRK-008': { id: 'MGPU-016', gpu: 'RTX PRO 6000 Blackwell Server Edition', vram: 96, precision: 'INT4', weight: 7, tp: 'Data Required', cmp: 'CMP-008', benchmark: 'BM-PROXY-RTX6K' },
  'MOD-003|WRK-009': { id: 'MGPU-017', gpu: 'B200 SXM6 180GB', vram: 180, precision: 'BF16', weight: 140, tp: 8, cmp: 'CMP-011', benchmark: 'BM-PROXY-B200' },
  'MOD-010|WRK-001': { id: 'MGPU-020', gpu: 'L40S', vram: 48, precision: 'Data Required', weight: 'Data Required', tp: 1, cmp: 'CMP-001', benchmark: 'BM-PROXY-L40S' },
  'MOD-009|WRK-011': { id: 'MGPU-023', gpu: 'RTX PRO 6000 Blackwell Server Edition', vram: 96, precision: 'INT8', weight: 24, tp: 1, cmp: 'CMP-008', benchmark: 'BM-PROXY-RTX6K' }
};
const hardware = {
  'CMP-001': { server: 'PowerEdge R760xa', gpus: 4, cpu: 'Intel Xeon Scalable 4th/5th Gen', sockets: 2, cores: 64, ram: 8, storage: 122.88, network: 'OCP 3.0 plus PCIe Gen5 expansion; exact NIC by configured BOM', cooling: 'Air', rack: 2, price: 71000 },
  'CMP-002': { server: 'PowerEdge R760xa', gpus: 4, cpu: 'Intel Xeon Scalable 4th/5th Gen', sockets: 2, cores: 64, ram: 8, storage: 122.88, network: 'OCP 3.0 plus PCIe Gen5 expansion; exact NIC by configured BOM', cooling: 'Air', rack: 2, price: 155000 },
  'CMP-003': { server: 'PowerEdge XE8640', gpus: 4, cpu: 'Intel Xeon Scalable 4th Gen', sockets: 2, cores: 56, ram: 'Data Required', storage: 'Data Required', network: 'High-speed GPU fabric; exact NIC by configured BOM', cooling: 'Liquid-cooled GPU tray / air-cooled CPU', rack: 4, price: 180000 },
  'CMP-005': { server: 'PowerEdge XE9680', gpus: 8, cpu: 'Intel Xeon Scalable 4th/5th Gen', sockets: 2, cores: 64, ram: 4, storage: 122.88, network: 'Up to 10 front-facing PCIe Gen5 slots; network options by accelerator configuration', cooling: 'Air / configured accelerator cooling', rack: 6, price: 410000 },
  'CMP-008': { server: 'PowerEdge XE7745', gpus: 8, cpu: 'AMD EPYC 9005', sockets: 2, cores: 192, ram: 3.072, storage: 122.88, network: '8 additional front-serviceable PCIe slots plus OCP 3.0', cooling: 'Air', rack: 4, price: 183000 },
  'CMP-009': { server: 'PowerEdge XE7745', gpus: 8, cpu: 'AMD EPYC 9005', sockets: 2, cores: 192, ram: 3.072, storage: 122.88, network: '8 additional front-serviceable PCIe slots plus OCP 3.0', cooling: 'Air', rack: 4, price: 375000 },
  'CMP-011': { server: 'PowerEdge XE9780', gpus: 8, cpu: 'Intel Xeon Scalable 6th Gen', sockets: 2, cores: 86, ram: 4, storage: 153.6, network: 'B300: 8 embedded CX8 OSFP; B200: up to 12 PCIe Gen5 x16 cards', cooling: 'Air', rack: 10, price: 500000 }
};
const benchmarks = {
  'BM-H100-5K500-C50': { type: 'Exact Published', confidence: 'High', ttft: 834.51, itl: 59.73, rpsGpu: .81437, sequencesGpu: 25, inputTpsGpu: 4071.85, outputTpsGpu: 407.185 },
  'BM-PROXY-L40S': { type: 'Nearest Planning Proxy', confidence: 'Low', ttft: 1000, itl: 50, rpsGpu: 1, sequencesGpu: 32, inputTpsGpu: 2000, outputTpsGpu: 500 },
  'BM-PROXY-L4': { type: 'Nearest Planning Proxy', confidence: 'Low', ttft: 1200, itl: 75, rpsGpu: .9, sequencesGpu: 16, inputTpsGpu: 900, outputTpsGpu: 180 },
  'BM-PROXY-H200': { type: 'Nearest Planning Proxy', confidence: 'Low', ttft: 1500, itl: 40, rpsGpu: .8, sequencesGpu: 32, inputTpsGpu: 12800, outputTpsGpu: 800 },
  'BM-PROXY-RTX6K': { type: 'Nearest Planning Proxy', confidence: 'Low', ttft: 1000, itl: 45, rpsGpu: 1.2, sequencesGpu: 32, inputTpsGpu: 4800, outputTpsGpu: 600 },
  'BM-PROXY-B200': { type: 'Nearest Planning Proxy', confidence: 'Low', ttft: 500, itl: 20, rpsGpu: 1.22, sequencesGpu: 32, inputTpsGpu: 5000, outputTpsGpu: 500 }
};
const $ = (selector) => document.querySelector(selector);
let bomOverrides = [null, null, null];
let selectedModelOverride = null;
async function loadCatalogManifest() {
  try {
    const response = await fetch('catalogs/manifest.json', { cache: 'no-store' });
    if (!response.ok) throw new Error('Catalog manifest unavailable');
    const manifest = await response.json();
    if (manifest.catalogVersion) $('#catalogBadge').textContent = `Catalog ${manifest.catalogVersion}`;
  } catch {
    $('#catalogBadge').textContent = 'Catalog snapshot';
  }
}
const n = (value, fallback = 0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
const roundup = (value) => Math.ceil(value);
const money = (value) => `$${Math.round(value).toLocaleString()}`;

function readInputs() { return Object.fromEntries(inputIds.map((id) => [id, document.querySelector(`#${id}`).value])); }
function applyInputs(data) { inputIds.forEach((id) => { if (data[id] !== undefined) document.querySelector(`#${id}`).value = data[id]; }); bomOverrides = Array.isArray(data.bomOverrides) ? data.bomOverrides : [null, null, null]; selectedModelOverride = data.selectedModelOverride || null; calculate(); }
function workloadRule(data, contextTier) {
  if (data.workloadType === 'Embedding Generation') return 'WRK-001';
  if (data.workloadType === 'LoRA / QLoRA Fine-Tuning') return 'WRK-008';
  if (data.workloadType === 'Full Fine-Tuning') return 'WRK-009';
  if (data.workloadType === 'Foundation Model Training') return 'WRK-010';
  if (data.workloadType === 'Computer Vision / Multimodal') return 'WRK-011';
  if (['Predictive ML', 'Forecasting'].includes(data.workloadType)) return 'WRK-012';
  if (data.workloadType === 'Batch Inference') return 'WRK-007';
  if (data.agenticComplexity === 'Advanced' || n(data.toolCalls) >= 5) return 'WRK-006';
  if (contextTier === 'CTX4') return 'WRK-005';
  if (data.queryComplexity === 'Complex' || data.responseDetail === 'Analytical') return 'WRK-004';
  if (data.queryComplexity === 'Medium' || data.responseDetail === 'Detailed' || n(data.toolCalls) >= 2) return 'WRK-003';
  return 'WRK-002';
}
function candidateModel(rule) { return { 'WRK-001': 'MOD-010', 'WRK-002': 'MOD-004', 'WRK-003': 'MOD-005', 'WRK-004': 'MOD-003', 'WRK-005': 'MOD-009', 'WRK-006': 'MOD-003', 'WRK-007': 'MOD-008', 'WRK-008': 'MOD-005', 'WRK-009': 'MOD-003', 'WRK-010': 'CUSTOM', 'WRK-011': 'MOD-009', 'WRK-012': 'NON-LLM' }[rule] || 'Data Required'; }
function eligibleModels(data, rule, requiredContext) {
  const mappedIds = [...new Set(Object.keys(mappings).filter((key) => key.endsWith(`|${rule}`)).map((key) => key.split('|')[0]))];
  return mappedIds.filter((modelId) => {
    const model = models[modelId];
    if (!model || (typeof model.context === 'number' && model.context < requiredContext)) return false;
    if (data.languageCoverage === 'Multilingual' && !model.languages.includes('Multilingual')) return false;
    if (data.languageCoverage === 'English + 2 languages' && !model.languages.includes('Multilingual') && model.languages.split(',').length < 3) return false;
    return true;
  });
}
function calculate() {
  const data = readInputs();
  const queryTokens = { Simple: 60, Medium: 200, Complex: 600 }[data.queryComplexity] || 60;
  const outputTokens = { Short: 200, Detailed: 600, Analytical: 1500, Concise: 200, Balanced: 500 }[data.responseDetail] || 600;
  const retainedTurns = { Single: 0, Short: 2, Medium: 5, Long: 10 }[data.conversationLength] || 5;
  const retrievedChunks = String(data.workloadType).includes('RAG') ? ({ Low: 2, Medium: 4, High: 8 }[data.retrievalDepth] || 0) : 0;
  const historyTokens = retainedTurns * (queryTokens + outputTokens + governed.formattingOverheadPerTurn);
  const retrievedContextTokens = retrievedChunks * governed.averageRetrievedChunkTokens;
  const toolResultTokens = n(data.toolCalls) * governed.averageToolResultTokens;
  const inputTokens = governed.systemPromptTokens + queryTokens + historyTokens + retrievedContextTokens + toolResultTokens;
  const requiredContext = inputTokens + outputTokens;
  const modelCallsPerSecond = n(data.requestsPerSecond, 5) * governed.validatedModelCallsPerRequest;
  const requiredInputTps = modelCallsPerSecond * inputTokens;
  const requiredOutputTps = modelCallsPerSecond * outputTokens;
  const activeSequences = n(data.concurrentSessions, 1000);
  const contextTier = requiredContext <= 8192 ? 'CTX1' : requiredContext <= 32768 ? 'CTX2' : requiredContext <= 65536 ? 'CTX3' : requiredContext <= 131072 ? 'CTX4' : 'Above 128K';
  const rule = workloadRule(data, contextTier);
  const automaticModelId = candidateModel(rule);
  const availableModelIds = eligibleModels(data, rule, requiredContext);
  const modelId = selectedModelOverride && availableModelIds.includes(selectedModelOverride) ? selectedModelOverride : (availableModelIds.includes(automaticModelId) ? automaticModelId : availableModelIds[0] || automaticModelId);
  const model = models[modelId] || { name: modelId, context: 'Data Required' };
  const mapping = mappings[`${modelId}|${rule}`];
  const benchmark = mapping ? benchmarks[mapping.benchmark] : null;
  const memoryWeight = mapping?.weight ?? 'Data Required';
  const runtimeWorkspace = memoryWeight === 'Data Required' ? 'Data Required' : Math.max(8, memoryWeight * .1);
  const kvCache = requiredContext * activeSequences * .00004;
  const totalVram = !mapping || [memoryWeight, runtimeWorkspace, kvCache].some((value) => value === 'Data Required') ? 'Data Required' : memoryWeight + runtimeWorkspace + kvCache;
  const memoryGpu = totalVram === 'Data Required' ? 'Data Required' : roundup(totalVram / (mapping.vram * .9));
  const ttftMatch = benchmark ? (benchmark.ttft <= n(data.p95Ttft, 1000) ? 'Pass' : 'Fail') : 'Data Required';
  const generationMatch = benchmark ? (1000 / benchmark.itl >= n(data.requiredGenerationSpeed, governed.requiredGenerationSpeed) ? 'Pass' : 'Fail') : 'Data Required';
  const throughputGpu = benchmark ? roundup(modelCallsPerSecond / benchmark.rpsGpu) : 'Data Required';
  const prefillGpu = benchmark ? roundup(requiredInputTps / benchmark.inputTpsGpu) : 'Data Required';
  const decodeGpu = benchmark ? roundup(requiredOutputTps / benchmark.outputTpsGpu) : 'Data Required';
  const concurrencyGpu = benchmark ? roundup(activeSequences / benchmark.sequencesGpu) : 'Data Required';
  const tpMinimum = mapping?.tp || 1;
  const industryBase = [memoryGpu, throughputGpu, prefillGpu, decodeGpu, concurrencyGpu, tpMinimum].some((value) => value === 'Data Required') ? 'Data Required' : Math.max(memoryGpu, throughputGpu, prefillGpu, decodeGpu, concurrencyGpu, tpMinimum);
  const server = mapping ? hardware[mapping.cmp] : null;
  const calculatedBaseValues = [1, 1.25, 1.5].map((factor) => industryBase === 'Data Required' || !server ? 'Data Required' : roundup(industryBase * factor / server.gpus) * server.gpus);
  const baseValues = calculatedBaseValues.map((value, index) => bomOverrides[index] === null || value === 'Data Required' ? value : Math.max(1, roundup(n(bomOverrides[index]))));
  const baseServers = baseValues.map((value) => value === 'Data Required' ? value : roundup(value / server.gpus));
  const growthServers = baseServers.map((value) => value === 'Data Required' ? value : roundup(value * Math.pow(1 + n(data.workloadGrowth) / 100, n(data.planningPeriod, 3))));
  const productionHa = growthServers.map((value) => value === 'Data Required' ? value : value + n(data.haServers));
  const drServers = productionHa.map((value) => value === 'Data Required' ? value : ['None', 'Backup Only'].includes(data.drRequirement) ? 0 : roundup(value * n(data.drCapacity, 50) / 100));
  const totalServers = productionHa.map((value, index) => value === 'Data Required' ? value : value + drServers[index]);
  const vram = baseValues.map((value) => value === 'Data Required' ? value : value * mapping.vram);
  const serverUnitPrice = n(data.serverUnitPrice, server?.price || 180000);
  const storageCostPerTb = n(data.storageCostPerTb, 1500);
  const nvmeCostPerTb = n(data.nvmeCostPerTb, 5000);
  const networkSwitchCost = n(data.networkSwitchCost, 50000);
  const supportRate = n(data.supportRate, 15) / 100;
  const contingencyRate = n(data.contingencyRate, 10) / 100;
  const hardwareTotals = totalServers.map((value) => value === 'Data Required' ? value : value * serverUnitPrice);
  const futureStorage = n(data.knowledgeBase, 10) * Math.pow(1 + n(data.dataGrowth) / 100, n(data.planningPeriod, 3));
  const hotNvme = String(data.workloadType).includes('RAG') ? futureStorage * .25 : 0;
  const switches = totalServers.map((value) => value === 'Data Required' ? value : Math.max(2, roundup(value * 2 / 32)));
  const storageCost = futureStorage * storageCostPerTb; const nvmeCost = hotNvme * nvmeCostPerTb; const networkCost = switches.map((value) => value === 'Data Required' ? value : value * networkSwitchCost);
  const support = hardwareTotals.map((value, i) => value === 'Data Required' ? value : (value + storageCost + nvmeCost + networkCost[i]) * supportRate);
  const capex = hardwareTotals.map((value, i) => value === 'Data Required' ? value : value + storageCost + nvmeCost + networkCost[i] + support[i]);
  const planningBudget = capex.map((value) => value === 'Data Required' ? value : value * (1 + contingencyRate));
  renderBom({ data, rule, model, mapping, benchmark, memoryGpu, ttftMatch, generationMatch, industryBase, baseValues, baseServers, growthServers, productionHa, drServers, totalServers, vram, hardwareTotals, server, serverUnitPrice, modelId, automaticModelId, availableModelIds });
  renderFinancials({ futureStorage, storageCost, hotNvme, nvmeCost, switches, networkCost, support, capex, planningBudget, hardwareTotals, contingencyRate });
  $('#budgetMinimum').textContent = typeof planningBudget[0] === 'number' ? money(planningBudget[0]) : planningBudget[0];
  $('#budgetModerate').textContent = typeof planningBudget[1] === 'number' ? money(planningBudget[1]) : planningBudget[1];
  $('#budgetHeavy').textContent = typeof planningBudget[2] === 'number' ? money(planningBudget[2]) : planningBudget[2];
  $('#growthLabel').textContent = `${n(data.planningPeriod, 3)}-year plan / ${n(data.workloadGrowth, 20)}% growth`;
  const tuned = [serverUnitPrice, storageCostPerTb, nvmeCostPerTb, networkSwitchCost, supportRate * 100, contingencyRate * 100].join('|') !== '180000|1500|5000|50000|15|10';
  $('#scenarioStatus').textContent = tuned ? 'User-tuned scenario' : 'Workbook baseline';
  $('#topProjectLabel').textContent = `${data.workloadType} sizing`; document.title = `${data.workloadType} sizing | InfraPilot AI`;
  return { ...data, bomOverrides, selectedModelOverride: selectedModelOverride === automaticModelId ? null : modelId, catalogVersion, catalogSource, calculated: { industryBase, calculatedBaseValues, baseValues, totalServers, planningBudget, rule, modelId }, updatedAt: new Date().toISOString() };
}
function renderBom(result) {
  const { data, rule, model, mapping, benchmark, memoryGpu, ttftMatch, generationMatch, industryBase, baseValues, baseServers, growthServers, productionHa, drServers, totalServers, vram, hardwareTotals, server, serverUnitPrice, modelId, automaticModelId, availableModelIds } = result;
  const modelSelect = $('#modelOverrideSelect');
  modelSelect.innerHTML = availableModelIds.map((id) => `<option value="${id}"${id === modelId ? ' selected' : ''}>${id} - ${models[id].name}</option>`).join('');
  $('#selectedModelLabel').textContent = model.name;
  $('#modelRuleLabel').textContent = `${rule} / ${modelId === automaticModelId ? 'automatic' : 'manual override'}`;
  $('#modelSelectionStatus').textContent = modelId === automaticModelId ? 'Automatically selected from workload, context, and language requirements.' : 'Manual model override active; hardware and BOM have been recalculated for this model.';
  const option = (values) => values;
  const readiness = benchmark?.type === 'Exact Published' ? 'Production Candidate' : 'Planning BOM';
  const rows = [
    ['BOM-001', 'Workload Rule', option([rule, rule, rule]), 'ID'], ['BOM-002', 'Selected Model', option([model.name, model.name, model.name]), 'model'], ['BOM-003', 'Benchmark Type', option([benchmark?.type || 'Benchmark Required', benchmark?.type || 'Benchmark Required', benchmark?.type || 'Benchmark Required']), 'benchmark'], ['BOM-004', 'GPU SKU', option([mapping?.gpu || 'Data Required', mapping?.gpu || 'Data Required', mapping?.gpu || 'Data Required']), 'GPU'], ['BOM-005', 'Industry Base GPU Count', [industryBase, industryBase, industryBase], 'GPUs'], ['BOM-006', 'Installed GPU Quantity', baseValues, 'GPUs'], ['BOM-007', 'Total Installed GPU VRAM', vram, 'GB'], ['BOM-008', 'Dell Server', [server?.server || 'Data Required', server?.server || 'Data Required', server?.server || 'Data Required'], 'server model'], ['BOM-009', 'Base Server Quantity', baseServers, 'servers'], ['BOM-010', 'Growth-Adjusted Servers', growthServers, 'servers'], ['BOM-011', 'Production + HA Servers', productionHa, 'servers'], ['BOM-012', 'DR Servers', drServers, 'servers'], ['BOM-013', 'Total Server Quantity', totalServers, 'servers'], ['BOM-014', 'CPU Platform', [server?.cpu || 'Data Required', server?.cpu || 'Data Required', server?.cpu || 'Data Required'], 'CPU platform'], ['BOM-015', 'CPU Sockets / Server', [server?.sockets || 'Data Required', server?.sockets || 'Data Required', server?.sockets || 'Data Required'], 'sockets/server'], ['BOM-016', 'Max Cores / Socket', [server?.cores || 'Data Required', server?.cores || 'Data Required', server?.cores || 'Data Required'], 'cores/socket'], ['BOM-017', 'Max System RAM / Server', [server?.ram || 'Data Required', server?.ram || 'Data Required', server?.ram || 'Data Required'], 'TB/server'], ['BOM-018', 'Max Local Storage / Server', [server?.storage || 'Data Required', server?.storage || 'Data Required', server?.storage || 'Data Required'], 'TB/server'], ['BOM-019', 'Network / Expansion', [server?.network || 'Data Required', server?.network || 'Data Required', server?.network || 'Data Required'], 'configuration'], ['BOM-020', 'Cooling', [server?.cooling || 'Data Required', server?.cooling || 'Data Required', server?.cooling || 'Data Required'], 'type'], ['BOM-021', 'Rack Units / Server', [server?.rack || 'Data Required', server?.rack || 'Data Required', server?.rack || 'Data Required'], 'U/server'], ['BOM-022', 'Budgetary Unit Price', [server ? money(server.price) : 'Data Required', server ? money(server.price) : 'Data Required', server ? money(server.price) : 'Data Required'], 'USD/server'], ['BOM-023', 'Budgetary Hardware Total', hardwareTotals.map((value) => typeof value === 'number' ? money(value) : value), 'USD'], ['BOM-024', 'Overall Readiness', [benchmark?.type === 'Exact Published' && ttftMatch === 'Pass' && generationMatch === 'Pass' ? 'Production Candidate' : 'Planning BOM', benchmark?.type === 'Exact Published' && ttftMatch === 'Pass' && generationMatch === 'Pass' ? 'Production Candidate' : 'Planning BOM', benchmark?.type === 'Exact Published' && ttftMatch === 'Pass' && generationMatch === 'Pass' ? 'Production Candidate' : 'Planning BOM'], 'status']
  ];
  rows[21][2] = [server ? money(serverUnitPrice) : 'Data Required', server ? money(serverUnitPrice) : 'Data Required', server ? money(serverUnitPrice) : 'Data Required'];
  rows[23][2] = [readiness, readiness, readiness];
  $('#bomBody').innerHTML = rows.map(([id, component, values, unit]) => `<tr><td>${id}</td><td>${component}${id === 'BOM-006' ? '<small class="editable-note">editable capacity</small>' : ''}</td>${values.map((value, index) => `<td>${id === 'BOM-006' ? `<input class="bom-override" data-option="${index}" type="number" min="1" step="1" value="${value}" aria-label="${component}, option ${index + 1}">` : value}</td>`).join('')}<td>${unit}</td></tr>`).join('');
  document.querySelectorAll('.bom-override').forEach((input) => input.addEventListener('input', () => { bomOverrides[Number(input.dataset.option)] = Math.max(1, roundup(n(input.value, 1))); calculate(); }));
}
function renderFinancials(result) {
  const rows = [['FIN-001', 'Dell GPU Server Hardware', result.capex.map((_, i) => result.hardwareTotals[i]), 'USD'], ['FIN-002', 'Future Capacity Storage', [result.futureStorage, result.futureStorage, result.futureStorage], 'TB'], ['FIN-003', 'Capacity Storage Cost', [result.storageCost, result.storageCost, result.storageCost], 'USD'], ['FIN-004', 'High-Speed RAG NVMe', [result.hotNvme, result.hotNvme, result.hotNvme], 'TB'], ['FIN-005', 'High-Speed NVMe Cost', [result.nvmeCost, result.nvmeCost, result.nvmeCost], 'USD'], ['FIN-006', 'Network Switches', result.switches, 'switches'], ['FIN-007', 'Network Cost', result.networkCost, 'USD'], ['FIN-008', 'Software and Support', result.support, 'USD'], ['FIN-009', 'Budgetary CAPEX', result.capex, 'USD'], ['FIN-010', 'Contingency', result.capex.map((value) => typeof value === 'number' ? value * result.contingencyRate : value), 'USD'], ['FIN-011', 'Total Planning Budget', result.planningBudget, 'USD']];
  $('#financialBody').innerHTML = rows.map(([id, component, values, unit]) => `<tr><td>${id}</td><td>${component}</td>${values.map((value) => `<td>${typeof value === 'number' ? (unit === 'USD' ? money(value) : value.toFixed(2)) : value}</td>`).join('')}<td>${unit}</td></tr>`).join('');
}
function getProjects() { try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); } catch { return []; } }
function saveProjects(projects) { localStorage.setItem(STORAGE_KEY, JSON.stringify(projects)); }
function showToast(message) { const toast = $('#toast'); toast.textContent = message; toast.classList.add('is-visible'); setTimeout(() => toast.classList.remove('is-visible'), 2400); }
function saveProject() { const project = calculate(); project.projectName = `${project.workloadType} sizing - ${new Date().toLocaleDateString()}`; const projects = getProjects().filter((item) => item.projectName !== project.projectName); projects.unshift(project); saveProjects(projects); localStorage.setItem(currentProjectKey, JSON.stringify(project)); $('#lastSavedLabel').textContent = `Saved ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`; showToast('Sizing project saved locally'); renderProjects(); }
function renderProjects() { const projects = getProjects(); const list = $('#projectList'); if (!projects.length) { list.innerHTML = '<div class="empty-state">No saved projects yet. Save your first sizing workspace to see it here.</div>'; return; } list.innerHTML = projects.map((project) => `<article class="project-card"><span class="section-kicker">${project.workloadType}</span><h3>${project.projectName}</h3><p>${project.calculated?.industryBase || '-'} base GPUs across the three BOM options.</p><div class="project-card-meta"><span>${new Date(project.updatedAt).toLocaleDateString()}</span><span>${project.calculated?.totalServers?.[1] || '-'} moderate servers</span></div><button class="button button-quiet load-project" data-project="${encodeURIComponent(JSON.stringify(project))}">Open project</button></article>`).join(''); list.querySelectorAll('.load-project').forEach((button) => button.addEventListener('click', () => { applyInputs(JSON.parse(decodeURIComponent(button.dataset.project))); showView('workspace'); showToast('Project loaded'); })); }
function showView(view) { document.querySelectorAll('.view').forEach((section) => section.classList.toggle('hidden', section.id !== `${view}View`)); document.querySelectorAll('.nav-item').forEach((button) => button.classList.toggle('is-active', button.dataset.view === view)); $('#pageTitle').textContent = view === 'workspace' ? 'AI infrastructure sizing.' : view === 'projects' ? 'Your infrastructure plans.' : view === 'catalogs' ? 'Manage sizing catalogs.' : 'Make the sizing logic legible.'; if (view === 'projects') renderProjects(); if (view === 'catalogs') renderCatalogRecords(); }
function getCatalogRecords() { try { return JSON.parse(localStorage.getItem(catalogRecordsKey) || '{}'); } catch { return {}; } }
function saveCatalogRecords(records) { localStorage.setItem(catalogRecordsKey, JSON.stringify(records)); }
function renderCatalogRecords() {
  const type = $('#catalogType').value; const records = getCatalogRecords()[type] || []; const list = $('#catalogRecordList');
  list.innerHTML = records.length ? records.map((record, index) => `<article class="catalog-record"><div><strong>${escapeHtml(record.id || record.ID || `Record ${index + 1}`)}</strong><span>${escapeHtml(record.name || record.model || record.server || record.GPU || 'Manual catalog record')}</span></div><button class="text-button delete-catalog-record" data-index="${index}">Remove</button><pre>${escapeHtml(JSON.stringify(record, null, 2))}</pre></article>`).join('') : '<div class="empty-state">No manual records for this catalog yet.</div>';
  list.querySelectorAll('.delete-catalog-record').forEach((button) => button.addEventListener('click', () => { const all = getCatalogRecords(); all[type].splice(Number(button.dataset.index), 1); saveCatalogRecords(all); renderCatalogRecords(); showToast('Manual catalog record removed'); }));
}
function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char])); }
function saveCatalogRecord() { const type = $('#catalogType').value; try { const record = JSON.parse($('#catalogRecordInput').value); if (!record || typeof record !== 'object' || Array.isArray(record)) throw new Error(); const all = getCatalogRecords(); all[type] = all[type] || []; all[type].push(record); saveCatalogRecords(all); $('#catalogRecordInput').value = ''; renderCatalogRecords(); showToast('Manual catalog record saved locally'); } catch { showToast('Enter one valid JSON object'); } }
function newProject() { $('#sizingForm').reset(); bomOverrides = [null, null, null]; selectedModelOverride = null; calculate(); $('#lastSavedLabel').textContent = 'Not saved yet'; showView('workspace'); }
$('#resetScenarioButton').addEventListener('click', () => { $('#serverUnitPrice').value = 180000; $('#storageCostPerTb').value = 1500; $('#nvmeCostPerTb').value = 5000; $('#networkSwitchCost').value = 50000; $('#supportRate').value = 15; $('#contingencyRate').value = 10; calculate(); });
$('#sizingForm').addEventListener('input', calculate); $('#sizingForm').addEventListener('change', calculate); $('#saveProjectButton').addEventListener('click', saveProject); $('#newProjectButton').addEventListener('click', newProject); $('#newProjectButtonLibrary').addEventListener('click', newProject); $('#resetButton').addEventListener('click', newProject);
document.querySelectorAll('.decision-grid input').forEach((input) => input.addEventListener('input', calculate));
$('#modelOverrideSelect').addEventListener('change', (event) => { selectedModelOverride = event.target.value; calculate(); });
$('#exportButton').addEventListener('click', () => { const blob = new Blob([JSON.stringify(getProjects(), null, 2)], { type: 'application/json' }); const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = 'infrapilot-ai-projects.json'; link.click(); URL.revokeObjectURL(link.href); });
$('#importInput').addEventListener('change', (event) => { const file = event.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { try { const imported = JSON.parse(reader.result); if (!Array.isArray(imported)) throw new Error(); saveProjects([...imported, ...getProjects()]); renderProjects(); showToast('Projects imported'); } catch { showToast('That file is not an InfraPilot AI project export'); } }; reader.readAsText(file); });
document.querySelectorAll('.nav-item').forEach((button) => button.addEventListener('click', () => showView(button.dataset.view)));
$('#catalogType').addEventListener('change', renderCatalogRecords); $('#saveCatalogRecord').addEventListener('click', saveCatalogRecord); $('#addCatalogRecord').addEventListener('click', () => $('#catalogRecordInput').focus());
$('#exportCatalogRecords').addEventListener('click', () => { const blob = new Blob([JSON.stringify(getCatalogRecords(), null, 2)], { type: 'application/json' }); const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = 'infrapilot-ai-manual-catalog-records.json'; link.click(); URL.revokeObjectURL(link.href); });
$('#importCatalogRecords').addEventListener('change', (event) => { const file = event.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { try { const imported = JSON.parse(reader.result); if (!imported || typeof imported !== 'object' || Array.isArray(imported)) throw new Error(); saveCatalogRecords({ ...getCatalogRecords(), ...imported }); renderCatalogRecords(); showToast('Manual catalog records imported'); } catch { showToast('That file is not a catalog export'); } }; reader.readAsText(file); });
renderCatalogRecords();
const current = localStorage.getItem(currentProjectKey); if (current) { try { const saved = JSON.parse(current); if (saved.workloadType) { applyInputs(saved); $('#lastSavedLabel').textContent = 'Current project loaded'; } else calculate(); } catch { calculate(); } } else calculate();
loadCatalogManifest();
