import { createPlugin, pkgRootFrom } from '@deepwhale-cn/dsh-guard-core/plugin';
import { registerDocxBatchApi } from './api.js';

const base = createPlugin({
  name: 'docx-batch',
  pkgRoot: pkgRootFrom(import.meta.url),
  skillFile: 'skill/docx-batch.md',
  binFile: 'bin/cli.mjs',
  description: 'Word 批量替换 / 合并 / 提取文本 / 页眉页脚 / 转 Markdown，一次处理整目录',
  whenToUse: '当用户要批量处理整目录的 Word（.docx）时使用：批量替换文本（合同改甲方、术语统一）、合并多份文档、提取正文到 txt、查看页眉页脚、转 Markdown。典型说法：「把这批 Word 里的旧公司名全换掉」「这些 docx 合并成一份」「Word 转 Markdown」。不用于 .doc 老格式、加密文档或需要保留复杂排版的场景。',
});

/** 技能名（与 cordis.patch.yml 的行 id 对应） */
export const name = base.name;
/** 技能服务 + 面板要用的 HTTP 出口 */
export const inject = [...base.inject, 'webServer'];

/**
 * 先注册技能（CLI/Agent 用），再把面板用的 HTTP 出口挂上。
 * webServer 缺失时只影响面板，CLI 与 skill 照常 —— 不让可选能力阻断启动。
 */
export function apply(ctx) {
  base.apply(ctx);
  try {
    if (ctx.webServer && typeof ctx.webServer.register === 'function') {
      registerDocxBatchApi(ctx);
      ctx.logger?.info?.('[docx-batch] 已挂面板出口 /dsh-docx-batch/api');
    }
  } catch (e) {
    ctx.logger?.warn?.('[docx-batch] 面板出口注册失败（CLI/skill 不受影响）: ' + (e && e.message ? e.message : e));
  }
}
