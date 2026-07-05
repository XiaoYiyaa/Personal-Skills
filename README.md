# Personal Skills

个人AI Agent技能集合。

## Skills

| Skill | 描述 |
|-------|------|
| [anki-generator](./anki-generator) | 题库转Anki导入文件，支持xlsx/toml/csv，自动生成cloze填空 |

## 安装

```bash
cp -r ./anki-generator ~/.codex/skills/
```

## anki-generator 使用

```bash
python3 scripts/convert.py --input 题库.xlsx --output output.txt --tag "第一章"
```

导入Anki时选择 **AwesomeSelect-3.x** 模板。

## License

MIT
