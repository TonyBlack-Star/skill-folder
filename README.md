# skill-folder

自建 Skill 集合仓库。每个技能独占一个子目录，结构参照 [JimLiu/baoyu-skills](https://github.com/JimLiu/baoyu-skills)。

## 目录结构

```
skills/
└── <skill-name>/
    ├── SKILL.md          # 技能定义（frontmatter + 说明 + 工作流）
    ├── scripts/          # 可执行脚本
    └── assets/           # 模板、资源文件
```

新增技能直接追加 `skills/<skill-name>/` 即可，互不干扰。

## 技能清单

| 技能 | 说明 |
|---|---|
| [shopify-traffic-sales-etl](skills/shopify-traffic-sales-etl/) | 整合 Shopify 落地页流量表与产品销售表（handle ↔ Product ID 桥接），生成多子表 XLSX、脱机 HTML 看板，支持按新日期增量更新，并输出结论导向的周度/月度分析报告（周口径：上周五–本周四） |
