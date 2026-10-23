import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';

export const { name, inject, apply } = createPlugin({
  name: 'docx-batch',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/docx-batch.md',
  binFile: 'bin/cli.mjs',
  description: 'Word 批量替换 / 合并 / 提取文本 / 页眉页脚 / 转 Markdown，一次处理整目录',
  whenToUse: '当用户要批量处理整目录的 Word（.docx）时使用：批量替换文本（合同改甲方、术语统一）、合并多份文档、提取正文到 txt、查看页眉页脚、转 Markdown。典型说法：「把这批 Word 里的旧公司名全换掉」「这些 docx 合并成一份」「Word 转 Markdown」。不用于 .doc 老格式、加密文档或需要保留复杂排版的场景。',
});
