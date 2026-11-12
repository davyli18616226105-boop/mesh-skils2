# 字段与计价契约

## 内部JSON

脚本需要UTF-8 JSON，数值可用十进制字符串，未知用null。仅支持确定性单价/已折后单价/按明确基数费率计算；累进、阶梯、税价转换及特殊结算不得套用普通单价模式。应先按真实口径完成相应公式，或留unknown。

```json
{
  "title":"客户产品报价草稿",
  "confirmed_business":false,
  "currency":"CNY",
  "commercial":{"tax_basis":null,"quote_date":null,"valid_until":null,"conditions_confirmed":false},
  "source_checks":[{"source_id":"source1","state":"matched"},{"source_id":"source2","state":"no_match"},{"source_id":"source3","state":"no_match"},{"source_id":"source4","state":"no_match"}],
  "rows":[
    {"id":"product1","group":"产品","product_line":"真实产品线","service":"真实服务项","selection":null,"functions":"有依据的功能及规格",
     "status":"unknown","mode":"unit","unit_price":null,"quantity":null,"unit":null,"discount":null,"periods":null,
     "source_id":"source1","source_ref":null,"sku_key":null,"price_current":false,"quote_authorized":false,"discount_authorized":false,"scope_confirmed":false,"basis_rows":[],"evidence":{}},
    {"id":"implementation1","group":"实施","status":"unknown"},
    {"id":"maintenance1","group":"运维","status":"unknown"}
  ]
}
```

## 模式与字段

- unit：行额=单价×数量×折扣×计费周期数。一次性或单期只有明确为1时填periods=1，不能脚本默认。
- net：unit_price是已折后单价，行额=该价×数量×周期，discount保持null，不再次折扣；D列注明已折后。
- percentage：basis_rows为前置且已确认的行ID，费率用rate（0—1），行额=明确基数行金额合计×费率×periods。单位、基数口径、服务等级/期限在功能清单说明；不能引用本行、后置行、含自身总价，避免循环。
- 阶梯计价：不直接支持。若档位尚未确认标unknown；若已确认需先开发/核验累进公式，不用全量×末档单价代替边际价。

## 状态

- confirmed：字段、服务范围与SKU已确认，必须source_id/source_ref/sku_key齐全，price_current、quote_authorized、discount_authorized、scope_confirmed为true，evidence逐字段包含来源或客户确认。unit/net须数量>0、periods>0，unit的discount在0—1且不为0；已确认为免费时用free。
- free / included：有当前适用的免费/套餐已含依据，amount=0；需要source_id、source_ref、sku_key、scope_confirmed、quote_authorized及evidence.scope/evidence.zero。备注写适用范围；不能把unknown伪装为free。
- not_selected / not_applicable：有客户明确不选或本次不适用依据，行额空。not_applicable亦须资料或本次已确认需求证据。
- unknown：资料、参数、价格或冲突未关闭，金额空，维度小计和最终总价阻断。参考字段允许有值，但未获授权/适用性未核验不计入正式金额。

source_checks必须覆盖四个固定source_id且不重复。只表示检索范围覆盖，不代表价格批准；blocked/conflict使confirmed_business无法通过，仍可出草稿。

## 金额与商务门禁

- 三个区块缺记录时各生成一条unknown空行。不删除实施/运维区块，也不自动标可选。
- 每区完整且金额/状态一致才出区小计；unknown不被SUM当0吞掉。未选或不适用是明确排除，不阻断其他合法行。
- 总价仅在confirmed_business=true、commercial.conditions_confirmed=true、税口径/报价日期/有效期与四源检查闭合、三维小计均有值时出现；否则空并写“待确认草稿”。
- commercial字段仅说明报价约定，脚本不计算税额、不替代审批；不同币种不能相加，当前脚本仅CNY。
- 内部核验输出包含逐行及三维参考复算、缺口、范围和证据，但不能直接附在对客表；纯文本写入需阻止公式注入，来源文本开头=、+、-、@一律作为文本，不作为Excel公式执行。
- 表中I列及汇总为真实公式；辅助输入放“报价条件”，只含非敏感状态、周期、费率、门禁信息，没有原始刊例或内部人名。

## 脚本与在线创建

1. Read脚本后，python3 技能目录/scripts/build_quote.py --input 实际JSON --output 实际新xlsx --report 实际内部报告JSON。
2. 仅模板：使用 --template --output 实际新xlsx，不提供--input；输出九列、三维、空值门禁与可填写条件，不写客户价格。
3. 调用当前sheet技能的+workbook-import：在文件所在授权cwd传相对文件名，等待ready=true、job_status=0，处理data.extra。
4. 按实际数量及业务读取该表验收公式、合并和缺口；不用导出替代验收。通过present_files交本次URL。
5. 脚本拒绝覆盖已有output/report。修改既有报价通过在线编辑流程，不能无声重生成覆盖。
