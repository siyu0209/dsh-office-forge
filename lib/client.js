window.__ModuleLoader__.load({id:`dsh-office-forge`,factory:require=>{var module={exports:{}};var exports=module.exports;let react=require("react");
const TITLE_STYLE={margin:0,fontSize:13,fontWeight:600,lineHeight:1.5};
const BODY_STYLE={margin:0,fontSize:12,lineHeight:1.7,opacity:.78};
const BUTTON_STYLE={font:"inherit",fontSize:12,padding:"4px 10px",border:"1px solid currentColor",borderRadius:4,background:"transparent",color:"inherit",cursor:"pointer",opacity:1};
const LIST_STYLE={margin:"6px 0 0 18px",padding:0,fontSize:12,lineHeight:1.6};
function checkLine(item){
  const status=typeof item.status===`string`?item.status:(item.ok?`ok`:`warn`);
  const detail=typeof item.detail===`string`?item.detail:``;
  const remedy=typeof item.remedy===`string`&&item.remedy?` — ${item.remedy}`:``;
  return `${status} · ${item.name||`check`}: ${detail}${remedy}`;
}
function OfficeForgeStatus(props){
  const subject=props&&props.subject?props.subject:{};
  const pkg=subject.pkg&&typeof subject.pkg===`object`?subject.pkg:{};
  const name=typeof pkg.name===`string`&&pkg.name?pkg.name:`dsh-office-forge`;
  const version=typeof pkg.version===`string`?pkg.version:``;
  const [state,setState]=react.useState({phase:`idle`,data:null,error:null});
  const runCheck=async()=>{
    setState({phase:`loading`,data:null,error:null});
    try{
      const response=await fetch(`/officeforge/doctor`,{headers:{accept:`application/json`}});
      const payload=await response.json().catch(()=>null);
      if(!response.ok)throw new Error(payload&&payload.error&&payload.error.message?payload.error.message:`HTTP ${response.status}`);
      setState({phase:`done`,data:payload,error:null});
    }catch(error){
      setState({phase:`error`,data:null,error:error instanceof Error?error.message:String(error)});
    }
  };
  const data=state.phase===`done`?state.data:null;
  const summary=data&&data.summary&&typeof data.summary===`object`?data.summary:null;
  const checks=data&&Array.isArray(data.checks)?data.checks:[];
  const rows=checks.map((item)=>react.createElement(`li`,{key:item.name||String(Math.random()),"data-officeforge-check-item":item.name||``},checkLine(item)));
  return react.createElement(`section`,{"data-officeforge-status":name},
    react.createElement(`p`,{style:TITLE_STYLE},version?`Office Forge ${version}`:`Office Forge`),
    react.createElement(`p`,{style:BODY_STYLE},`Word / Excel / PowerPoint / PDF 文件级工具。点“检查环境”查看 Python、PDF、OCR、COM、LibreOffice 与宿主 officeKit 的可用状态；模块与功能开关在 profile 的 cordis.patch.yml 中配置。`),
    react.createElement(`div`,{style:{display:`flex`,gap:`8px`,alignItems:`center`,margin:`8px 0`,flexWrap:`wrap`}},
      react.createElement(`button`,{type:`button`,"data-officeforge-check":``,disabled:state.phase===`loading`,onClick:runCheck,style:BUTTON_STYLE},state.phase===`loading`?`检查中…`:`检查环境`),
      summary?react.createElement(`span`,{style:{fontSize:12}},`health=${data.health||`unknown`} · ok=${summary.ok??0} warn=${summary.warn??0} missing=${summary.missing??0}`):null
    ),
    state.phase===`error`?react.createElement(`p`,{style:{margin:0,fontSize:12,color:`#c0392b`}},`检查失败：${state.error}`):null,
    rows.length?react.createElement(`ul`,{style:LIST_STYLE,"data-officeforge-check-list":``},rows):null,
    react.createElement(`p`,{style:BODY_STYLE},`本卡片只读：不会修改配置，也不会在插件页执行任何文档操作。`)
  );
}
function apply(ctx){
  ctx.slots.inject(`plugins.detail.section`,()=>ctx.slots.register({name:`plugins.detail.section`,id:`officeforge.status`,order:40},OfficeForgeStatus));
}
const inject=[`slots`];
exports.apply=apply;exports.inject=inject;return module.exports}});
