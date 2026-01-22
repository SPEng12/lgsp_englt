const { useState, useMemo, useEffect, useRef } = React;
const { 
    BarChart, Bar, ComposedChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
    PieChart, Pie, Cell, Legend, Treemap 
} = window.Recharts || {};

// --- Icons Component ---
const LucideIcon = ({ name, size = 18, className = "", ...props }) => {
    if (!window.lucide) return null;
    const iconData = window.lucide.icons[name];
    if (!iconData) return null;
    return React.createElement('svg', { ...props, width: size, height: size, stroke: "currentColor", strokeWidth: 2, fill: "none", strokeLinecap: "round", strokeLinejoin: "round", className, viewBox: "0 0 24 24", xmlns: "http://www.w3.org/2000/svg" }, iconData.map(([tag, attrs], index) => React.createElement(tag, { ...attrs, key: index })));
};

const Icons = {
    Upload: (p) => <LucideIcon name="Upload" {...p} />,
    FileSpreadsheet: (p) => <LucideIcon name="FileSpreadsheet" {...p} />,
    Building2: (p) => <LucideIcon name="Building2" {...p} />,
    TrendingUp: (p) => <LucideIcon name="TrendingUp" {...p} />,
    TrendingDown: (p) => <LucideIcon name="TrendingDown" {...p} />,
    DollarSign: (p) => <LucideIcon name="DollarSign" {...p} />,
    Zap: (p) => <LucideIcon name="Zap" {...p} />,
    LayoutDashboard: (p) => <LucideIcon name="LayoutDashboard" {...p} />,
    Table: (p) => <LucideIcon name="Table" {...p} />,
    Leaf: (p) => <LucideIcon name="Leaf" {...p} />,
    Users: (p) => <LucideIcon name="Users" {...p} />,
    Activity: (p) => <LucideIcon name="Activity" {...p} />,
    ArrowRight: (p) => <LucideIcon name="ArrowRight" {...p} />,
    BarChart2: (p) => <LucideIcon name="BarChart2" {...p} />,
    ChevronLeft: (p) => <LucideIcon name="ChevronLeft" {...p} />,
    ChevronRight: (p) => <LucideIcon name="ChevronRight" {...p} />,
    Sun: (p) => <LucideIcon name="Sun" {...p} />,
    Moon: (p) => <LucideIcon name="Moon" {...p} />,
    Download: (p) => <LucideIcon name="Download" {...p} />,
    PieChart: (p) => <LucideIcon name="PieChart" {...p} />,
    AlertTriangle: (p) => <LucideIcon name="AlertTriangle" {...p} />,
    CheckCircle: (p) => <LucideIcon name="CheckCircle" {...p} />,
    FolderOpen: (p) => <LucideIcon name="FolderOpen" {...p} />,
    CheckSquare: (p) => <LucideIcon name="CheckSquare" {...p} />,
    Square: (p) => <LucideIcon name="Square" {...p} />,
    Map: (p) => <LucideIcon name="Map" {...p} />,
    Calendar: (p) => <LucideIcon name="Calendar" {...p} />,
    Layers: (p) => <LucideIcon name="Layers" {...p} />,
};

// --- Data & Constants ---
const BUILDING_LIST = ["DP2", "DP3_CA", "LGC_DP3", "LGC_D22", "CNS_D22", "CNS_D25", "D22_CA", "D25_CA", "ISC", "SLC", "LGD", "LXH", "LGIT", "LGES", "LGCS", "LGHH", "LGU", "LXG", "DP2_2단계"];

const TENANT_DISPLAY_MAPPING = { 
    "LGE": "LG전자", "LGES": "LG에너지솔루션", 
    "LGD": "LG디스플레이", 
    "LGIT": "LG이노텍", 
    "LGC_D22": "LG화학_D22",
    "LGC": "LG화학", 
    "LGHH": "LG생활건강", 
    "LGU": "LG유플러스", 
    "CNS": "LG씨앤에스", "LGCS": "LG화학_E5,E7", 
    "LXH": "LX하우시스", 
    "LXG": "LX글라스", 
    "LGE_DP3": "LG전자_DP3", 
    "DP3_CA": "DP3_공용", "D22_CA": "D22_공용", "D25_CA": "D25_공용" 
};

const TENANT_GROUPS = {
    "DP3": ["LG디스플레이", "LG에너지솔루션", "LG이노텍", "LG화학", "LX글라스", "LX하우시스", "LG전자_DP3"],
    "D22/D25": ["LG생활건강", "LG씨앤에스", "LG유플러스", "LG화학_D22"]
};

const MONTHS = ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"];
const SOURCES = ["전체", "전력", "도시가스", "중온수", "상하수도", "재이용수"];
const DATA_TYPES = [{ id: 'Cost', label: '비용', unit: '원' }, { id: 'Usage', label: '사용량', unit: 'Usage' }, { id: 'TOE', label: '에너지', unit: 'TOE' }, { id: 'tCO2', label: '온실가스', unit: 'tCO2' }];

// Improved Color Palette
const COLORS = { 
    prev: "#94a3b8", // Slate 400
    curr: "#881337", // Dark Rose
    currHighlight: "#e11d48", // Bright Rose (Latest)
    projected: "#fbbf24", 
    mix: ["#A50034", "#F97316", "#3B82F6", "#10B981", "#6366F1"] 
};

// --- Utils ---
const getUnit = (dataType, source) => {
    if (dataType === 'Cost') return '원';
    if (dataType === 'TOE') return 'TOE';
    if (dataType === 'tCO2') return 'tCO2';
    if (source === '전력') return 'kWh';
    if (source === '중온수') return 'MWh';
    if (['도시가스', '상하수도', '재이용수'].includes(source)) return 'm³';
    return 'TOE';
};

const formatValue = (val, dataType) => {
    if (val === null || val === undefined) return "-";
    if (dataType === 'Cost') {
        return new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 0 }).format(val);
    }
    const digits = Math.abs(val) < 1000 ? 1 : 0;
    return new Intl.NumberFormat('ko-KR', { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(val);
};

const AnimatedNumber = ({ value, formatter, className }) => {
    const [displayValue, setDisplayValue] = useState(0);
    useEffect(() => {
        let start = 0; const end = value; const duration = 800; const startTime = performance.now();
        const animate = (currentTime) => {
            const elapsed = currentTime - startTime; const progress = Math.min(elapsed / duration, 1);
            const ease = 1 - Math.pow(1 - progress, 3);
            setDisplayValue(start + (end - start) * ease);
            if (progress < 1) requestAnimationFrame(animate);
        };
        requestAnimationFrame(animate);
    }, [value]);
    return <span className={className}>{formatter(displayValue)}</span>;
};

const CustomizedTreemapContent = (props) => {
    const { root, depth, x, y, width, height, index, payload, colors, rank, name, value, rate } = props;
    if (width < 30 || height < 30) return null;
    let bgColor = "#64748b"; 
    if (rate > 0.1) bgColor = "#780026";
    else if (rate > 0) bgColor = "#A50034";
    else if (rate < -0.05) bgColor = "#059669";
    else if (rate < 0) bgColor = "#0d9488";

    return (
        <g>
            <rect x={x} y={y} width={width} height={height} style={{ fill: bgColor, stroke: '#fff', strokeWidth: 2 / (depth + 1e-10), strokeOpacity: 1 / (depth + 1e-10), rx: 4, ry: 4 }} />
            {width > 50 && height > 30 && (
                <foreignObject x={x} y={y} width={width} height={height}>
                    <div xmlns="http://www.w3.org/1999/xhtml" style={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', color: '#fff', overflow: 'hidden', padding: '2px' }}>
                        <div style={{ fontSize: '11px', fontWeight: 'bold', textAlign: 'center', textShadow: '0 1px 2px rgba(0,0,0,0.3)', lineHeight: '1.1', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth:'100%' }}>{name}</div>
                        <div style={{ fontSize: '9px', marginTop: '1px', opacity: 0.9 }}>
                            {rate > 0 ? '▲' : (rate < 0 ? '▼' : '-')}{Math.abs(rate * 100).toFixed(1)}%
                        </div>
                    </div>
                </foreignObject>
            )}
        </g>
    );
};

// --- Main Dashboard ---
function EnergyDashboard() {
    const [buildingData, setBuildingData] = useState(null);
    const [tenantData, setTenantData] = useState(null);
    const [viewMode, setViewMode] = useState('Tenant');
    const [selectedEntities, setSelectedEntities] = useState([]);
    const [selectedDataType, setSelectedDataType] = useState('Cost');
    const [selectedSource, setSelectedSource] = useState('전체');
    const [isFileUploaded, setIsFileUploaded] = useState(false);
    const [loading, setLoading] = useState(false);
    const [loadingMessage, setLoadingMessage] = useState("");
    const [isLibLoaded, setIsLibLoaded] = useState(false);
    const [isSidebarOpen, setIsSidebarOpen] = useState(true);
    const [isDarkMode, setIsDarkMode] = useState(false);
    
    // Dynamic Year State
    const [years, setYears] = useState({ prev: 2024, curr: 2025 });

    useEffect(() => { const i = setInterval(() => { if (window.Recharts && window.XLSX && window.lucide) { setIsLibLoaded(true); clearInterval(i); } }, 500); return () => clearInterval(i); }, []);
    useEffect(() => { document.documentElement.classList.toggle('dark', isDarkMode); }, [isDarkMode]);
    
    // --- Excel Processing Logic ---
    const processExcelData = (arrayBuffer) => {
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        const bData = {}, tData = {};
        const init = (t, n) => { if (!t[n]) { t[n] = {}; DATA_TYPES.forEach(dt => { t[n][dt.id] = {}; SOURCES.forEach(s => t[n][dt.id][s] = MONTHS.map(m => ({ month: m, prev: 0, curr: 0 }))); }); } };
        BUILDING_LIST.forEach(b => init(bData, b));
        
        // Sheet finding logic
        let sheetName = wb.SheetNames.find(n => (n.includes("실적") || n.includes("에너지")) && (n.includes("년") || n.includes("20")));
        if (!sheetName) sheetName = wb.SheetNames[0];

        let detectedPrev = new Date().getFullYear() - 1;
        let detectedCurr = new Date().getFullYear();

        if (sheetName) {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
            
            // 1. 헤더(1월, 2월 등) 찾기 - 검색 범위를 200행까지 확장하여 B88 주변도 커버
            let hIdx = -1, mIdxs = [];
            for (let i = 0; i < Math.min(rows.length, 200); i++) {
                if (rows[i] && rows[i].findIndex(c => String(c).trim() === "1월") > -1) { 
                    hIdx = i; 
                    for(let m = 0; m < 12; m++) mIdxs.push(rows[i].findIndex(c => String(c).trim() === "1월") + m); 
                    break; 
                }
            }

            // 2. 년도 감지 로직 수정 (B열 기준 검색)
            const yearMatches = new Set();
            rows.forEach((row) => {
                // 행이 있고, B열(index 1)에 데이터가 있는 경우
                if (row && row.length > 1) {
                    const bColValue = String(row[1] || "").trim();
                    // 20XX 패턴 찾기
                    const matches = bColValue.match(/20\d{2}/);
                    if (matches) {
                        yearMatches.add(parseInt(matches[0]));
                    }
                }
            });

            // B열에서 년도를 못 찾았을 경우를 대비한 Fallback (행 전체 스캔)
            if (yearMatches.size < 2) {
                 rows.forEach(row => {
                     const str = (row || []).join(" ");
                     const matches = str.match(/20\d{2}/g);
                     if(matches) matches.forEach(y => yearMatches.add(parseInt(y)));
                 });
            }
            
            if(yearMatches.size >= 2) {
                const sortedYears = Array.from(yearMatches).sort((a,b) => a - b);
                detectedPrev = sortedYears[sortedYears.length - 2];
                detectedCurr = sortedYears[sortedYears.length - 1];
            }
            
            if (hIdx > -1) {
                for (let i = hIdx + 1; i < rows.length; i++) {
                    const row = rows[i]; if (!row || !row.length) continue;
                    const meta = row.slice(0, 10).map(c => String(c || "")).join(" ");
                    
                    let targetKey = 'curr';
                    if (meta.includes(String(detectedPrev)) || meta.includes(String(detectedPrev).slice(2) + "년")) targetKey = 'prev';
                    else if (meta.includes(String(detectedCurr)) || meta.includes(String(detectedCurr).slice(2) + "년")) targetKey = 'curr';
                    
                    const rFull = row.map(c => String(c || "").toLowerCase()).join("|");
                    let dT = null; 
                    if(rFull.includes("비용")) dT='Cost'; 
                    else if(rFull.includes("사용량")&&(rFull.includes("kwh")||rFull.includes("m3")||rFull.includes("㎥")||rFull.includes("mwh"))) dT='Usage'; 
                    else if(rFull.includes("toe")) dT='TOE'; 
                    else if(rFull.includes("tco2") || rFull.includes("온실가스")) dT='tCO2';
                    if (!dT) continue;

                    let src = "기타"; 
                    if(rFull.includes("전력")) src="전력"; 
                    else if(rFull.includes("가스")||rFull.includes("lng")) src="도시가스"; 
                    else if(rFull.includes("상하수도")||rFull.includes("수도")) src="상하수도"; 
                    else if(rFull.includes("중온수")||rFull.includes("난방")) src="중온수"; 
                    else if(rFull.includes("재이용수")) src="재이용수";
                    
                    const process = (target) => {
                        mIdxs.forEach((c, m) => {
                            let v = row[c]; if(typeof v==='string') v=parseFloat(v.replace(/,/g,''));
                            if(typeof v==='number' && !isNaN(v)) { 
                                target[dT][src][m][targetKey] += v; 
                                target[dT]['전체'][m][targetKey] += v; 
                            }
                        });
                    };
                    
                    if (!rFull.includes("입주사")) {
                        const mCells = row.slice(0, 10).map(c => String(c || "").trim());
                        const mStr = mCells.join(" ");
                        let matchedB = BUILDING_LIST.find(b => {
                            if (b==="DP2_2단계") return mStr.includes("DP2_2단계") || (mStr.includes("DP2")&&mStr.includes("2단계")) || mStr.includes("2차부지");
                            if (b==="DP2") return (mCells.some(c=>c==="DP2"||(c.includes("DP2")&&!c.includes("DP2_2단계"))) && !(mStr.includes("2단계")||mStr.includes("2차")));
                            return mCells.some(c=>c===b || (c.includes(b)&&!b.includes("DP2")));
                        });
                        if (matchedB) process(bData[matchedB]);
                    }
                    if (rFull.includes("입주사")) {
                        let matchedT = null;
                        for (const c of row.slice(4, 12)) {
                            let v = String(c || "").trim(); if(!v||v==="사용량"||v==="비용") continue;
                            if (v==="LGE_DP3") v="LG전자_DP3"; if (v==="LGE") matchedT="LG에너지솔루션";
                            else { 
                                const keys = Object.keys(TENANT_DISPLAY_MAPPING).sort((a,b) => b.length - a.length);
                                const k = keys.find(k => v===k || v.startsWith(k));
                                matchedT = k ? TENANT_DISPLAY_MAPPING[k] : (Object.values(TENANT_DISPLAY_MAPPING).includes(v) ? v : (["LG","LX","CNS"].some(p=>v.startsWith(p))?v:null)); 
                            }
                            if (matchedT) break;
                        }
                        if (matchedT) { init(tData, matchedT); process(tData[matchedT]); }
                    }
                }
            }
        }
        
        [bData, tData].forEach(d => Object.values(d).forEach(dt => Object.values(dt).forEach(src => Object.values(src).forEach(arr => {
            let last = -1; arr.forEach((x,i)=>{if(x.curr>0)last=i;});
            let sPrev=0, sCurr=0; if(last>-1) for(let k=0;k<=last;k++){sPrev+=arr[k].prev||0; sCurr+=arr[k].curr||0;}
            let rate = sPrev>0 ? (sCurr-sPrev)/sPrev : 0;
            arr.forEach((x,i)=>{ 
                if(i>last) { x.curr=null; x.projected=(x.prev||0)*(1+rate); } else x.projected=null;
                x.diff = (x.curr!==null&&x.prev!==null) ? x.curr-x.prev : 0;
            });
        }))));

        setYears({ prev: detectedPrev, curr: detectedCurr });
        setBuildingData(bData); setTenantData(tData); 
        setViewMode('Tenant'); 
        setSelectedEntities([]);
        setIsFileUploaded(true);
    };

    // --- Auto Load Data Logic ---
    useEffect(() => {
        const loadDefaultData = async () => {
            if (!isLibLoaded) return;
            setLoading(true);
            setLoadingMessage("데이터 파일 탐색 중 (data 폴더 확인)...");
            const candidates = ['./data/energy_data.xlsx', './data/data.xlsx', './data/energy.xlsx'];
            let loaded = false;
            for (const path of candidates) {
                try {
                    const response = await fetch(path);
                    if (response.ok) {
                        const arrayBuffer = await response.arrayBuffer();
                        processExcelData(arrayBuffer);
                        loaded = true;
                        setLoadingMessage("");
                        break;
                    }
                } catch (e) { continue; }
            }
            if (!loaded) {
                console.log("자동 로드 실패: data 폴더에 파일이 없거나 접근할 수 없습니다. (수동 업로드 대기)");
                setLoadingMessage("");
            }
            setLoading(false);
        };
        loadDefaultData();
    }, [isLibLoaded]);

    const handleFileUpload = (e) => {
        e.preventDefault(); if (!isLibLoaded) return;
        let file = e.dataTransfer ? e.dataTransfer.files[0] : e.target.files[0]; if (!file) return;
        setLoading(true); setLoadingMessage("파일 분석 중...");
        const reader = new FileReader();
        reader.onload = (evt) => {
            try { processExcelData(evt.target.result); } 
            catch(e) { console.error(e); alert("파일 처리 중 오류가 발생했습니다."); } 
            finally { setLoading(false); setLoadingMessage(""); }
        };
        reader.readAsArrayBuffer(file);
    };

    const toggleEntity = (name) => {
        setSelectedEntities(prev => prev.includes(name) ? prev.filter(e => e !== name) : [...prev, name]);
    };

    const toggleAllEntities = () => {
        const allKeys = Object.keys(activeData || {}).sort();
        setSelectedEntities(selectedEntities.length === allKeys.length ? [] : allKeys);
    }

    const selectGroup = (groupName) => {
        const groupMembers = TENANT_GROUPS[groupName];
        if (!groupMembers) return;
        const availableMembers = groupMembers.filter(m => activeData && activeData[m]);
        const allSelected = availableMembers.every(m => selectedEntities.includes(m));
        if (allSelected) {
            setSelectedEntities(prev => prev.filter(e => !availableMembers.includes(e)));
        } else {
            const newSet = new Set([...selectedEntities, ...availableMembers]);
            setSelectedEntities(Array.from(newSet).sort());
        }
    };

    const activeData = viewMode === 'Building' ? buildingData : tenantData;
    
    const aggregatedData = useMemo(() => {
        const emptyAgg = MONTHS.map(m => ({ month: m, prev: 0, curr: 0, projected: 0, count: 0 }));
        if (!activeData || selectedEntities.length === 0) return emptyAgg;
        const agg = MONTHS.map((m, idx) => ({ month: m, prev: 0, curr: 0, projected: 0, count: 0 }));
        selectedEntities.forEach(entity => {
            const data = activeData[entity];
            if (!data) return;
            let targetArr = (selectedDataType === 'Usage' && selectedSource === '전체') ? data['TOE']['전체'] : data[selectedDataType][selectedSource];
            if (targetArr) {
                targetArr.forEach((d, idx) => {
                    agg[idx].prev += (d.prev || 0);
                    agg[idx].curr += (d.curr || 0);
                    agg[idx].projected += (d.projected || 0);
                });
            }
        });
        return agg;
    }, [activeData, selectedEntities, selectedDataType, selectedSource]);

    const currentData = aggregatedData;

    // Detect Latest Month Index for Chart Highlighting
    const latestMonthIndex = useMemo(() => {
        let idx = -1;
        currentData.forEach((d, i) => {
            if (d.curr !== null && d.curr > 0) idx = i;
        });
        return idx;
    }, [currentData]);

    const mixData = useMemo(() => {
        if (!activeData || selectedEntities.length === 0 || selectedSource !== '전체') return null;
        const sourceKeys = ["전력", "도시가스", "중온수", "상하수도", "재이용수"];
        const result = sourceKeys.map(key => ({ name: key, value: 0 }));
        selectedEntities.forEach(entity => {
            const data = activeData[entity];
            if(!data) return;
            sourceKeys.forEach((key, idx) => {
                const arr = data[selectedDataType][key];
                const val = arr.reduce((a,c) => a + (c.curr || c.projected || 0), 0);
                result[idx].value += val;
            });
        });
        return result.filter(d => d.value > 0);
    }, [activeData, selectedEntities, selectedDataType, selectedSource]);

    const summary = useMemo(() => {
        if (!currentData || !Array.isArray(currentData) || !currentData.length) return { totalPrev:0, totalCurr:0, diff:0, percent:0, projection:0, latestVal: 0, latestMonth: '-', latestDiff: 0, latestPercent: 0 };
        
        const tPrev = currentData.reduce((a,c) => a + (c.prev||0), 0);
        const tCurr_actual = currentData.reduce((a,c) => a + (c.curr||0), 0);
        const tCurr_proj_total = tCurr_actual + currentData.reduce((a,c) => a + (c.projected||0), 0);

        let lastIdx = latestMonthIndex;
        const tPrev_ytd = currentData.slice(0, lastIdx+1).reduce((a,c) => a + (c.prev||0), 0);
        
        const latestObj = lastIdx > -1 ? currentData[lastIdx] : null;
        const latestVal = latestObj ? latestObj.curr : 0;
        const latestPrev = latestObj ? latestObj.prev : 0;
        const latestMonth = latestObj ? latestObj.month : '-';
        const latestDiff = latestVal - latestPrev;
        const latestPercent = latestPrev ? (latestDiff / latestPrev) * 100 : 0;

        return { 
            totalPrev: tPrev_ytd, 
            totalCurr: tCurr_actual, 
            diff: tCurr_actual - tPrev_ytd, 
            percent: tPrev_ytd ? (tCurr_actual - tPrev_ytd)/tPrev_ytd*100 : 0, 
            projection: tCurr_proj_total,
            latestVal,
            latestMonth,
            latestDiff,
            latestPercent
        };
    }, [currentData, latestMonthIndex]);
    
    const treemapData = useMemo(() => {
        if (!activeData || selectedEntities.length === 0) return [];
        const list = [];
        selectedEntities.forEach(key => {
            const data = activeData[key];
            if (!data) return;
            let totalPrev = 0;
            let totalCurr_proj = 0;
            let targetArr = (selectedDataType === 'Usage' && selectedSource === '전체') ? data['TOE']['전체'] : data[selectedDataType][selectedSource];
            targetArr.forEach(d => {
                totalPrev += (d.prev || 0);
                totalCurr_proj += (d.curr || 0) + (d.projected || 0);
            });
            if (totalCurr_proj > 0) {
                const rate = totalPrev ? (totalCurr_proj - totalPrev) / totalPrev : 0;
                list.push({ 
                    name: viewMode === 'Building' ? key : (TENANT_DISPLAY_MAPPING[key] || key), 
                    size: totalCurr_proj, 
                    rate: rate 
                });
            }
        });
        return list.sort((a,b) => b.size - a.size);
    }, [activeData, selectedEntities, selectedDataType, selectedSource, viewMode]);

    const currentUnit = useMemo(() => getUnit(selectedDataType, selectedSource), [selectedDataType, selectedSource]);
    const entityList = useMemo(() => activeData ? Object.keys(activeData).sort() : [], [activeData]);

    const exportToExcel = () => {
        const safeCurrentData = Array.isArray(currentData) ? currentData : [];
        const wsData = [
            ["구분", ...MONTHS, "합계"],
            [`${years.prev}년`, ...safeCurrentData.map(d => d.prev), summary.totalPrev],
            [`${years.curr}년`, ...safeCurrentData.map(d => d.curr || d.projected), summary.projection]
        ];
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");
        XLSX.writeFile(wb, `Energy_Pro_Report_${selectedDataType}.xlsx`);
    };

    if (!isFileUploaded) {
        return (
            <div className="min-h-screen bg-slate-50 dark:bg-slate-950 flex flex-col items-center justify-center p-4">
                <div className="max-w-xl w-full bg-white dark:bg-slate-900 rounded-[2rem] shadow-xl p-10 text-center border border-slate-100 dark:border-slate-800 animate-fade-in-up">
                    <div className="w-20 h-20 bg-gradient-to-tr from-brand-red to-rose-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-glow animate-bounce">
                        <Icons.FileSpreadsheet className="w-10 h-10 text-white" />
                    </div>
                    <h1 className="text-3xl font-black mb-3 text-slate-800 dark:text-white">LGSP 에너지레터</h1>
                    <p className="text-sm text-slate-500 mb-8">
                    </p>
                    <div className="border-3 border-dashed border-slate-200 dark:border-slate-700 rounded-2xl p-10 hover:border-brand-red hover:bg-rose-50/10 transition-all cursor-pointer group relative"
                        onDragOver={e=>e.preventDefault()} onDrop={handleFileUpload}>
                        <input type="file" accept=".xlsx,.xls" className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" onChange={handleFileUpload} disabled={!isLibLoaded} />
                        <Icons.Upload className="w-12 h-12 mx-auto mb-4 text-slate-300 group-hover:text-brand-red transition-colors" />
                        <p className="text-base font-bold text-slate-400 group-hover:text-brand-red transition-colors">{isLibLoaded?"파일 업로드":"시스템 로딩 중..."}</p>
                    </div>
                    {loading && <div className="mt-6 text-brand-red font-bold animate-pulse">{loadingMessage || "데이터 분석 중..."}</div>}
                </div>
            </div>
        );
    }

    return (
        <div className="flex h-screen bg-slate-100 dark:bg-slate-950 font-sans overflow-hidden text-slate-700 dark:text-slate-300">
            <aside className={`${isSidebarOpen ? 'w-64' : 'w-20'} bg-white dark:bg-slate-900 shadow-xl z-20 flex flex-col border-r border-slate-200 dark:border-slate-800 transition-all duration-300`}>
                <div className="p-6 pb-4 flex flex-col gap-6">
                    <div className="flex items-center gap-3 overflow-hidden">
                        <div className="min-w-[40px] w-10 h-10 bg-gradient-to-br from-brand-red to-rose-600 rounded-xl flex items-center justify-center text-white shadow-glow">
                            <Icons.Activity className="w-6 h-6" />
                        </div>
                        {isSidebarOpen && (
                            <div className="animate-fade-in-up">
                                <h1 className="text-xl font-black leading-none text-slate-800 dark:text-white">LGSP 에너지레터</h1>
                                <p className="text-[10px] text-slate-400 font-bold mt-0.5 tracking-widest">beta</p>
                            </div>
                        )}
                    </div>
                    {isSidebarOpen ? (
                        <div className="space-y-3">
                            <div className="flex p-1 bg-slate-100 dark:bg-slate-800 rounded-xl">
                                <button onClick={() => { setViewMode('Tenant'); setSelectedEntities([]); }} className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${viewMode === 'Tenant' ? 'bg-white dark:bg-slate-700 text-brand-red shadow-sm' : 'text-slate-500'}`}>입주사</button>
                                <button onClick={() => { setViewMode('Building'); setSelectedEntities([]); }} className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all ${viewMode === 'Building' ? 'bg-white dark:bg-slate-700 text-brand-red shadow-sm' : 'text-slate-500'}`}>건물</button>
                            </div>
                            <button onClick={toggleAllEntities} className="w-full py-2 text-xs font-bold border border-slate-200 dark:border-slate-700 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors flex items-center justify-center gap-2">
                                {selectedEntities.length === entityList.length ? <Icons.CheckSquare className="w-4 h-4 text-brand-red"/> : <Icons.Square className="w-4 h-4"/>}
                                {selectedEntities.length === entityList.length ? "전체 해제" : "전체 선택"}
                            </button>
                            {viewMode === 'Tenant' && (
                                <div className="grid grid-cols-2 gap-2">
                                    {Object.keys(TENANT_GROUPS).map(gName => (
                                        <button key={gName} onClick={() => selectGroup(gName)} className="py-2 text-[10px] font-bold border border-slate-200 dark:border-slate-700 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors flex items-center justify-center gap-1">
                                            <Icons.Layers className="w-3 h-3"/>
                                            {gName}
                                        </button>
                                    ))}
                                </div>
                            )}
                        </div>
                    ) : (
                        <button onClick={() => setIsSidebarOpen(true)} className="p-2 mx-auto bg-slate-100 dark:bg-slate-800 rounded-lg text-brand-red"><Icons.Users className="w-5 h-5"/></button>
                    )}
                </div>
                <div className="flex-1 overflow-y-auto px-4 pb-4 custom-scrollbar space-y-1">
                    {entityList.map(name => {
                        const isSelected = selectedEntities.includes(name);
                        return (
                            <button key={name} onClick={() => toggleEntity(name)}
                                className={`w-full text-left px-4 py-3 rounded-xl text-sm font-semibold transition-all flex items-center gap-3 group border ${isSelected ? 'bg-gradient-to-r from-brand-red to-rose-600 text-white shadow-md border-transparent' : 'bg-white dark:bg-slate-900 border-transparent hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-500'}`}>
                                {isSelected ? <Icons.CheckSquare className="w-4 h-4 shrink-0"/> : <Icons.Square className="w-4 h-4 shrink-0 opacity-50"/>}
                                {isSidebarOpen && <span className="truncate">{viewMode === 'Building' ? name : (TENANT_DISPLAY_MAPPING[name] || name)}</span>}
                            </button>
                        );
                    })}
                </div>
                <div className="p-4 border-t border-slate-200 dark:border-slate-800">
                      <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className="w-full py-2 flex items-center justify-center text-slate-400 hover:text-brand-red transition-colors">
                        {isSidebarOpen ? <Icons.ChevronLeft className="w-5 h-5" /> : <Icons.ChevronRight className="w-5 h-5" />}
                    </button>
                </div>
            </aside>

            <main className="flex-1 flex flex-col overflow-hidden relative">
                <header className="bg-white/90 dark:bg-slate-900/90 backdrop-blur-md border-b border-slate-200 dark:border-slate-800 h-16 flex items-center justify-between px-8 sticky top-0 z-30 shadow-sm">
                    <div className="flex items-center gap-3 min-w-fit">
                        <h2 className="text-xl font-black text-slate-800 dark:text-white tracking-tight flex items-center gap-3">
                            {selectedEntities.length > 1 ? `${selectedEntities.length}개 항목 선택됨` : (selectedEntities.length === 1 ? (TENANT_DISPLAY_MAPPING[selectedEntities[0]] || selectedEntities[0]) : "선택 안됨")}
                            <span className="bg-rose-50 dark:bg-rose-900/30 text-brand-red dark:text-rose-400 text-[10px] px-2.5 py-0.5 rounded-full font-bold uppercase tracking-widest border border-rose-100 dark:border-rose-800">
                                {viewMode === 'Building' ? '건물' : '입주사'}
                            </span>
                        </h2>
                    </div>
                    
                    <div className="flex-1 flex items-center justify-center gap-4 overflow-x-auto px-4 hide-scrollbar">
                        <div className="hidden md:flex bg-slate-100 dark:bg-slate-800 p-1 rounded-full border border-slate-200 dark:border-slate-700 shrink-0">
                            {SOURCES.map(src => (
                                <button key={src} onClick={() => setSelectedSource(src)} className={`px-4 py-1.5 text-xs font-bold rounded-full transition-all ${selectedSource === src ? 'bg-white dark:bg-slate-600 text-slate-900 dark:text-white shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>{src}</button>
                            ))}
                        </div>
                        <div className="flex gap-2 shrink-0">
                            {DATA_TYPES.map(type => (
                                <button key={type.id} onClick={() => setSelectedDataType(type.id)} className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all flex items-center gap-1.5 ${selectedDataType === type.id ? 'bg-slate-800 dark:bg-slate-100 text-white dark:text-slate-900 shadow-lg scale-105' : 'bg-white dark:bg-slate-800 text-slate-500 hover:border-slate-300'}`}>
                                    {type.label.split(' ')[0]}
                                </button>
                            ))}
                        </div>
                    </div>

                    <div className="flex items-center justify-end gap-3 min-w-fit">
                        <button onClick={exportToExcel} className="flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-slate-600 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 hover:text-brand-red transition-all shadow-sm">
                            <Icons.Download className="w-4 h-4" />
                            엑셀 다운로드
                        </button>
                        <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 rounded-full bg-white dark:bg-slate-800 border border-slate-200 text-slate-500 hover:text-brand-red shadow-sm transition-transform hover:rotate-12">
                            {isDarkMode ? <Icons.Sun className="w-5 h-5" /> : <Icons.Moon className="w-5 h-5" />}
                        </button>
                    </div>
                </header>

                <div className="flex-1 overflow-y-auto p-6 custom-scrollbar bg-slate-50/50 dark:bg-slate-950">
                    <div className="max-w-[1600px] mx-auto space-y-6 pb-20">
                        
                        {/* Smart Insight Banner */}
                        <div className={`w-full p-4 rounded-xl border flex items-start gap-3 animate-fade-in-up ${summary.percent > 5 ? 'bg-orange-50 border-orange-200 text-orange-800' : 'bg-blue-50 border-blue-200 text-blue-800'}`}>
                            {summary.percent > 5 ? <Icons.AlertTriangle className="w-5 h-5 shrink-0 mt-0.5" /> : <Icons.CheckCircle className="w-5 h-5 shrink-0 mt-0.5" />}
                            <div>
                                <h4 className="font-bold text-sm">Smart Insight (선택 합계)</h4>
                                <p className="text-xs opacity-80 mt-1">
                                    {selectedEntities.length > 0 ? (
                                        <>
                                            선택된 항목들의 {DATA_TYPES.find(d=>d.id===selectedDataType).label}이(가) 전년 동기 대비 <span className="font-bold">{Math.abs(summary.percent).toFixed(1)}% {summary.percent > 0 ? '증가' : '감소'}</span>했습니다. 
                                            {summary.percent > 10 ? ' 급격한 증가 추세이므로 원인 파악 및 관리가 필요합니다.' : ' 안정적인 관리 상태를 보이고 있습니다.'}
                                        </>
                                    ) : (
                                        "좌측 사이드바에서 분석할 대상을 선택해주세요."
                                    )}
                                </p>
                            </div>
                        </div>

                        {/* KPI Cards */}
                        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 animate-fade-in-up delay-100">
                            {/* Card 1 */}
                            <div className="relative overflow-hidden rounded-[2rem] bg-slate-900 dark:bg-slate-800 text-white shadow-xl group card-hover">
                                <div className="absolute inset-0 bg-gradient-to-br from-slate-800 to-slate-950 opacity-90"></div>
                                <div className="absolute -right-10 -top-10 w-40 h-40 bg-brand-red rounded-full blur-[60px] opacity-40"></div>
                                <div className="relative z-10 p-5 flex flex-col justify-between h-full">
                                    <div className="flex justify-between items-start">
                                        <div className="text-sm font-bold text-slate-300 opacity-80">{selectedSource} 전체 합계</div>
                                        <div className={`px-2 py-1 rounded text-[10px] font-bold ${summary.percent > 0 ? 'bg-red-500/20 text-red-300' : 'bg-blue-500/20 text-blue-300'}`}>
                                            {summary.percent > 0 ? '▲' : '▼'} {Math.abs(summary.percent).toFixed(1)}% YoY
                                        </div>
                                    </div>
                                    <div className="mt-4">
                                        <h3 className="text-3xl font-black tracking-tight"><AnimatedNumber value={summary.totalCurr} formatter={(v) => formatValue(v, selectedDataType)} /></h3>
                                        <p className="text-xs text-slate-400 font-medium mt-1">단위: {currentUnit}</p>
                                    </div>
                                </div>
                            </div>

                            {/* Card 2: Current Month Metric (Highlighted) */}
                            <div className="rounded-[2rem] bg-emerald-50/50 dark:bg-emerald-900/10 p-5 border border-emerald-200 dark:border-emerald-800 shadow-lg ring-2 ring-emerald-500/20 card-hover group relative overflow-hidden">
                                <div className="absolute top-0 right-0 p-4 opacity-10">
                                    <Icons.Calendar className="w-20 h-20 text-emerald-500" />
                                </div>
                                <div className="flex justify-between items-center mb-4 relative z-10">
                                    <h3 className="text-sm font-bold text-emerald-800 dark:text-emerald-200">당월({summary.latestMonth}) 실적</h3>
                                    <div className="p-2 bg-white dark:bg-emerald-800/30 rounded-lg text-emerald-600 dark:text-emerald-300"><Icons.Calendar className="w-5 h-5" /></div>
                                </div>
                                <h3 className="text-3xl font-black text-slate-800 dark:text-white relative z-10"><AnimatedNumber value={summary.latestVal} formatter={(v) => formatValue(v, selectedDataType)} /></h3>
                                <p className={`text-xs mt-2 font-bold flex items-center gap-1 relative z-10 ${summary.latestPercent > 0 ? 'text-brand-red' : 'text-emerald-600'}`}>
                                    {summary.latestPercent > 0 ? <Icons.TrendingUp className="w-3 h-3"/> : <Icons.TrendingDown className="w-3 h-3"/>}
                                    {summary.latestPercent > 0 ? '+' : ''}{summary.latestPercent.toFixed(1)}% (전년 동월 대비)
                                </p>
                            </div>

                            {/* Card 3: Projection */}
                            <div className="rounded-[2rem] bg-white dark:bg-slate-800 p-5 border border-slate-100 dark:border-slate-700 shadow-card card-hover group">
                                <div className="flex justify-between items-center mb-4">
                                    <h3 className="text-sm font-bold text-slate-500">연말 예상</h3>
                                    <div className="p-2 bg-blue-50 dark:bg-blue-900/20 rounded-lg text-blue-500"><Icons.TrendingUp className="w-5 h-5" /></div>
                                </div>
                                <h3 className="text-3xl font-black text-slate-800 dark:text-white"><AnimatedNumber value={summary.projection} formatter={(v) => formatValue(v, selectedDataType)} /></h3>
                                <div className="w-full bg-slate-100 rounded-full h-1.5 mt-4 overflow-hidden">
                                    <div className="bg-blue-500 h-full rounded-full" style={{width: '75%'}}></div>
                                </div>
                                <p className="text-xs text-slate-400 mt-2 text-right">진척률 (예상)</p>
                            </div>

                            {/* Card 4: Gap Analysis */}
                            <div className="rounded-[2rem] bg-white dark:bg-slate-800 p-5 border border-slate-100 dark:border-slate-700 shadow-card card-hover">
                                <div className="flex justify-between items-center mb-4">
                                    <h3 className="text-sm font-bold text-slate-500">실적 차이 (Gap)</h3>
                                    <div className="p-2 bg-orange-50 dark:bg-orange-900/20 rounded-lg text-orange-500"><Icons.Activity className="w-5 h-5" /></div>
                                </div>
                                <h3 className={`text-3xl font-black ${summary.diff > 0 ? 'text-brand-red' : 'text-blue-500'}`}>
                                    {summary.diff > 0 ? '+' : ''}<AnimatedNumber value={summary.diff} formatter={(v) => formatValue(v, selectedDataType)} />
                                </h3>
                                <p className="text-xs text-slate-400 mt-2">전년 대비 누적 차이</p>
                            </div>
                        </div>

                        {/* Charts Row */}
                        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 animate-fade-in-up delay-200">
                            {/* Monthly Trend Chart (2/3 width) */}
                            <div className="lg:col-span-2 rounded-[2rem] bg-white dark:bg-slate-800 p-6 shadow-card border border-slate-100 dark:border-slate-700">
                                <div className="flex justify-between items-center mb-6">
                                    <h3 className="font-bold text-lg text-slate-800 dark:text-white flex items-center gap-2">
                                        <Icons.BarChart2 className="w-5 h-5 text-slate-400"/> 월별 추세 분석
                                    </h3>
                                    <div className="flex gap-3 text-xs font-bold">
                                        <div className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full bg-slate-400"></span> {years.prev}년</div>
                                        <div className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full bg-[#881337]"></span> {years.curr}년</div>
                                        <div className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full bg-brand-rose shadow-lg shadow-rose-500/50"></span> 당월(강조)</div>
                                        <div className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full bg-amber-400"></span> 예상</div>
                                    </div>
                                </div>
                                <div className="h-[300px]">
                                    <ResponsiveContainer width="100%" height="100%">
                                        <ComposedChart data={currentData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }} barGap={0}>
                                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={isDarkMode ? "#334155" : "#f1f5f9"} />
                                            <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize:11}} dy={10} />
                                            <YAxis axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize:11}} tickFormatter={(val)=>formatValue(val,selectedDataType)} width={60} />
                                            <Tooltip formatter={(val)=>formatValue(val,selectedDataType)} contentStyle={{borderRadius:'12px', border:'none', boxShadow:'0 10px 15px -3px rgba(0, 0, 0, 0.1)'}} cursor={{fill:'transparent'}} />
                                            
                                            {/* Previous Year Area */}
                                            <Area type="monotone" dataKey="prev" name={`${years.prev}년`} stroke={COLORS.prev} fill={COLORS.prev} fillOpacity={0.2} strokeWidth={2} dot={false} />
                                            
                                            {/* Current Year Bar with Highlight Logic */}
                                            <Bar dataKey="curr" name={`${years.curr}년`} radius={[4,4,0,0]} barSize={24}>
                                                {currentData.map((entry, index) => {
                                                    const isLatest = index === latestMonthIndex;
                                                    return (
                                                        <Cell 
                                                            key={`cell-${index}`} 
                                                            fill={isLatest ? COLORS.currHighlight : COLORS.curr} 
                                                            stroke={isLatest ? '#fca5a5' : 'none'}
                                                            strokeWidth={isLatest ? 2 : 0}
                                                            style={isLatest ? { filter: 'drop-shadow(0px 0px 4px rgba(225, 29, 72, 0.6))' } : {}}
                                                        />
                                                    );
                                                })}
                                            </Bar>
                                            
                                            {/* Projected Bar */}
                                            <Bar dataKey="projected" name="예상" fill={COLORS.projected} radius={[4,4,0,0]} barSize={24} />
                                        </ComposedChart>
                                    </ResponsiveContainer>
                                </div>
                            </div>

                            {/* Energy Portfolio (Pie Chart) */}
                            <div className="rounded-[2rem] bg-white dark:bg-slate-800 p-6 shadow-card border border-slate-100 dark:border-slate-700 flex flex-col">
                                <h3 className="font-bold text-lg text-slate-800 dark:text-white flex items-center gap-2 mb-4">
                                    <Icons.PieChart className="w-5 h-5 text-slate-400"/> 에너지 포트폴리오
                                </h3>
                                {selectedSource === '전체' && mixData && mixData.length > 0 ? (
                                    <div className="flex-1 min-h-[250px] relative">
                                        <ResponsiveContainer width="100%" height="100%">
                                            <PieChart>
                                                <Pie data={mixData} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                                                    {mixData.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS.mix[index % COLORS.mix.length]} />)}
                                                </Pie>
                                                <Tooltip formatter={(val)=>formatValue(val,selectedDataType)} />
                                                <Legend verticalAlign="bottom" height={36} iconType="circle" />
                                            </PieChart>
                                        </ResponsiveContainer>
                                        <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
                                            <div className="text-center">
                                                <p className="text-xs text-slate-400 font-bold">TOTAL</p>
                                                <p className="text-lg font-black text-slate-700 dark:text-slate-200">{formatValue(mixData.reduce((a,c)=>a+c.value,0), selectedDataType)}</p>
                                            </div>
                                        </div>
                                    </div>
                                ) : (
                                    <div className="flex-1 flex items-center justify-center text-slate-400 text-sm bg-slate-50 dark:bg-slate-900 rounded-xl border border-dashed border-slate-200 text-center p-4">
                                        {selectedSource === '전체' ? '데이터가 없습니다.' : "'전체' 보기 모드에서만\n포트폴리오가 제공됩니다."}
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* Energy Map (Treemap) */}
                        <div className="w-full p-6 rounded-[2rem] bg-white dark:bg-slate-800 shadow-card border border-slate-100 dark:border-slate-700 animate-fade-in-up">
                            <div className="flex justify-between items-center mb-4">
                                <h3 className="font-bold text-lg text-slate-800 dark:text-white flex items-center gap-2">
                                    <Icons.Map className="w-5 h-5 text-slate-400"/> Energy Map (선택 항목)
                                </h3>
                                <div className="flex gap-2 text-[10px] font-bold text-white">
                                    <span className="px-2 py-1 rounded bg-[#780026]">급증 (&gt;10%)</span>
                                    <span className="px-2 py-1 rounded bg-[#A50034]">증가</span>
                                    <span className="px-2 py-1 rounded bg-[#64748b]">보합</span>
                                    <span className="px-2 py-1 rounded bg-[#059669]">감소</span>
                                </div>
                            </div>
                            <div className="h-[400px] w-full">
                                {treemapData.length > 0 ? (
                                    <ResponsiveContainer width="100%" height="100%">
                                        <Treemap
                                            data={treemapData}
                                            dataKey="size"
                                            aspectRatio={4 / 3}
                                            stroke="#fff"
                                            fill="#8884d8"
                                            isAnimationActive={false}
                                            animationDuration={0}
                                            content={<CustomizedTreemapContent />}
                                        >
                                            <Tooltip 
                                                content={({ payload }) => {
                                                    if (payload && payload.length) {
                                                        const d = payload[0].payload;
                                                        return (
                                                            <div className="bg-white dark:bg-slate-900 p-3 rounded-xl shadow-xl border border-slate-100 dark:border-slate-700">
                                                                <p className="font-bold text-slate-800 dark:text-white mb-1">{d.name}</p>
                                                                <p className="text-xs text-slate-500">규모: {formatValue(d.size, selectedDataType)}</p>
                                                                <p className={`text-xs font-bold ${d.rate > 0 ? 'text-brand-red' : 'text-emerald-600'}`}>
                                                                    전년비: {d.rate > 0 ? '+' : ''}{(d.rate * 100).toFixed(1)}%
                                                                </p>
                                                            </div>
                                                        );
                                                    }
                                                    return null;
                                                }}
                                            />
                                        </Treemap>
                                    </ResponsiveContainer>
                                ) : (
                                    <div className="flex items-center justify-center h-full text-slate-400 bg-slate-50 dark:bg-slate-900 rounded-xl border border-dashed border-slate-200">
                                        좌측 사이드바에서 분석할 대상을 선택하면 Map이 표시됩니다.
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* Detailed Table */}
                        <div className="rounded-[2rem] bg-white dark:bg-slate-800 shadow-card border border-slate-100 dark:border-slate-700 animate-fade-in-up delay-300">
                            <div className="p-5 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center bg-slate-50/50 rounded-t-[2rem]">
                                <h3 className="text-lg font-bold text-slate-800 dark:text-white flex items-center gap-2">
                                    <Icons.Table className="w-5 h-5 text-slate-400"/> 상세 내역 (합계 기준)
                                </h3>
                                <span className="text-[10px] text-slate-400 font-medium">* 단위: {currentUnit}</span>
                            </div>
                            
                            <div className="p-4 w-full overflow-hidden">
                                <div className="w-full">
                                    <table className="w-full text-right table-fixed border-collapse compact-table">
                                        <thead>
                                            <tr className="text-[10px] md:text-xs text-slate-500 border-b border-slate-200 dark:border-slate-700">
                                                <th className="py-2 px-1 text-left font-bold w-[8%] bg-slate-50 dark:bg-slate-900">구분</th>
                                                {MONTHS.map(m => (
                                                    <th key={m} className="py-2 px-1 font-medium w-[7%] text-center text-slate-400">{m.replace('월','')}</th>
                                                ))}
                                                <th className="py-2 px-1 font-bold text-slate-700 dark:text-slate-300 w-[8%] bg-slate-50/50 dark:bg-slate-800/50">Total</th>
                                            </tr>
                                        </thead>
                                        <tbody className="text-[10px] md:text-xs font-medium text-slate-600 dark:text-slate-400 divide-y divide-slate-100 dark:divide-slate-800">
                                            {/* Prev Year Row */}
                                            <tr className="group hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
                                                <td className="py-2 px-1 text-left font-bold text-slate-500">{years.prev}</td>
                                                {currentData.map((d, i) => (
                                                    <td key={i} className="py-2 px-1 tabular-nums tracking-tight text-center opacity-80 group-hover:opacity-100">
                                                        {formatValue(d.prev, selectedDataType)}
                                                    </td>
                                                ))}
                                                <td className="py-2 px-1 font-bold text-slate-800 dark:text-slate-200 bg-slate-50/30 dark:bg-slate-800/30">
                                                    {formatValue(summary.totalPrev, selectedDataType)}
                                                </td>
                                            </tr>
                                            
                                            {/* Curr Year Row */}
                                            <tr className="group hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
                                                <td className="py-2 px-1 text-left font-black text-brand-red">{years.curr}</td>
                                                {currentData.map((d, i) => {
                                                    const val = d.curr !== null ? d.curr : d.projected;
                                                    const isProj = d.curr === null;
                                                    const prev = d.prev || 0;
                                                    const diffRate = prev ? ((val - prev) / prev) : 0;
                                                    
                                                    let cellClass = "text-slate-800 dark:text-slate-200";
                                                    if (!isProj && prev > 0) {
                                                        if (diffRate > 0.1) cellClass = "text-brand-red font-bold bg-red-50/50 dark:bg-red-900/10 rounded-sm";
                                                        else if (diffRate < -0.05) cellClass = "text-emerald-600 font-bold bg-emerald-50/50 dark:bg-emerald-900/10 rounded-sm";
                                                    }
                                                    if (isProj) cellClass = "text-amber-500 italic";

                                                    return (
                                                        <td key={i} className="py-2 px-1 text-center">
                                                            <div className={`py-0.5 px-1 ${cellClass} tabular-nums tracking-tight inline-block w-full`}>
                                                                {formatValue(val, selectedDataType)}
                                                            </div>
                                                        </td>
                                                    )
                                                })}
                                                <td className="py-2 px-1 font-black text-brand-red bg-slate-50/30 dark:bg-slate-800/30">
                                                    {formatValue(summary.projection, selectedDataType)}
                                                </td>
                                            </tr>
                                            
                                            {/* Gap Row */}
                                            <tr className="group hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors border-t border-slate-200 dark:border-slate-700 border-dashed">
                                                <td className="py-2 px-1 text-left font-bold text-slate-400 text-[9px]">Gap</td>
                                                {currentData.map((d, i) => {
                                                    const val = d.curr !== null ? d.curr : d.projected;
                                                    const diff = val - (d.prev || 0);
                                                    return (
                                                        <td key={i} className={`py-2 px-1 text-center text-[9px] tabular-nums tracking-tight ${diff > 0 ? 'text-brand-red' : 'text-emerald-500'}`}>
                                                            {diff > 0 ? '+' : ''}{formatValue(diff, selectedDataType)}
                                                        </td>
                                                    )
                                                })}
                                                <td className={`py-2 px-1 text-[9px] font-bold text-right ${summary.diff > 0 ? 'text-brand-red' : 'text-emerald-500'}`}>
                                                    {summary.diff > 0 ? '+' : ''}{formatValue(summary.diff, selectedDataType)}
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                    </div>
                </div>
            </main>
        </div>
    );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<EnergyDashboard />);