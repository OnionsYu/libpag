'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { before, after, describe, it } = require('node:test');
const { PNG } = require('pngjs');
const { launchBrowser } = require('../../dist/lib/browser-engine');
const { takeSnapshot } = require('../../dist/lib/browser-snapshot');

const binary = process.env.PAGX_HTML_TEST_BIN ||
  path.resolve(__dirname, '../../../../cmake-build-debug/pagx');

describe('snapshot ownership preserves native paint while reducing layers', {
  skip: !fs.existsSync(binary),
}, () => {
  let browser, page;
  before(async () => {
    ({ browser } = await launchBrowser());
    page = await browser.newPage();
    await page.setViewport({ width: 320, height: 240, deviceScaleFactor: 1 });
  });
  after(async () => { if (browser) await browser.close(); });

  async function compare(content, frames) {
    await page.setContent(`<html><head><style>
      *{box-sizing:border-box}html,body{margin:0;padding:0}
      body{width:320px;height:240px;background:white;font:18px Arial}
      </style></head><body>${content}</body></html>`);
    const snapshot = await page.evaluate(takeSnapshot);
    assert.equal(snapshot.scene.version, 1);
    const owner = snapshot.scene.sources.find(source => source.elementId === 'owner');
    const authored = snapshot.scene.sources.find(source => source.elementId === 'authored');
    assert.equal(authored.parentId, owner.id);
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pagx-layer-hierarchy-'));
    function run(args) {
      const result = spawnSync(binary, args, { encoding: 'utf8', timeout: 30000 });
      assert.equal(result.status, 0, `${result.stderr}; ${dir}`);
    }
    const files = {};
    // The reference imports identical layout/text/effect values as independent layers.
    // Removing only ownership annotations disables primitive lowering, keeping a pixel oracle
    // independent of the native representation selected for generated content.
    for (const [name, html] of [
      ['layers', snapshot.html.replace(/\sdata-pagx-node-(?:role|source|parent)="[^"]*"/g, '')],
      ['contents', snapshot.html],
    ]) {
      const input = path.join(dir, `${name}.html`);
      files[name] = path.join(dir, `${name}.pagx`);
      fs.writeFileSync(input, html);
      run(['import', input, '--html-no-snapshot', '-o', files[name]]);
    }
    const original = fs.readFileSync(files.layers, 'utf8');
    const normalized = fs.readFileSync(files.contents, 'utf8');
    const count = xml => (xml.match(/<Layer\b/g) || []).length;
    assert.ok(count(normalized) < count(original), `no layer reduction; ${dir}`);
    for (const name of ['owner', 'authored']) {
      assert.equal((normalized.match(new RegExp(`name="${name}"`, 'g')) || []).length, 1,
        `authored layer ${name} must remain unique; ${dir}`);
    }
    assert.equal((normalized.match(/<Text\b/g) || []).length,
      (original.match(/<Text\b/g) || []).length, `editable text was lost; ${dir}`);
    for (const { scale = 1, time = 0 } of frames) {
      const images = [];
      for (const name of ['layers', 'contents']) {
        const output = path.join(dir, `${name}-${scale}-${time}.png`);
        run(['render', files[name], '-o', output, '--scale', String(scale), '--time', String(time)]);
        images.push(PNG.sync.read(fs.readFileSync(output)));
      }
      assert.equal(images[0].width, images[1].width);
      assert.equal(images[0].height, images[1].height);
      assert.ok(images[0].data.equals(images[1].data),
        `paint changed at scale=${scale}, time=${time}; ${dir}`);
    }
    fs.rmSync(dir, { recursive: true });
  }

  it('keeps text, gradients, padding, clips and fractional transforms at multiple scales', async () => {
    await compare(`<div id="owner" name="owner" style="position:relative;left:13.4px;top:7.6px;
      width:285px;height:210px;border:1.4px solid #369;padding:5.7px;background:#eee">
      <div style="position:absolute;left:8.4px;top:5.6px;width:240px;height:35px;
        line-height:23.4px">Fractional text 123</div>
      <div id="authored" name="authored" style="position:absolute;left:9px;top:55px;width:230px;height:42px;
        background:linear-gradient(90deg,red,blue);background-clip:text;color:transparent">
        Gradient <b>text</b></div>
      <div style="position:absolute;left:12px;top:110px;width:210px;height:50px;
        transform:scale(.833333);transform-origin:0 0;overflow:hidden;border-radius:12px;
        background:#fcd">Transformed clipping</div>
      </div>`, [{ scale: 1 }, { scale: 1.5 }, { scale: 2 }]);
  });

  it('keeps generated text inside its animated source group throughout playback', async () => {
    await compare(`<style id="__pagx_anim_keyframes">
      @keyframes pagxAnim0{from{opacity:.3;transform:translate(0,0) scale(.8)}
        to{opacity:1;transform:translate(23px,11px) scale(1.1)}}
      </style><div id="owner" name="owner" style="position:relative;width:240px;height:180px;
        animation:pagxAnim0 2s linear both;background:#cef">
      <div id="authored" name="authored" style="position:absolute;left:11px;top:17px;width:180px;height:35px">
        Animated text</div>
      <div style="position:absolute;left:25px;top:65px;width:180px;height:35px">Second line</div>
      </div>`, [{ time: 0 }, { time: .5 }, { time: 1 }, { time: 1.75 }]);
  });

  it('keeps a sole flex text item centered at fractional sizes', async () => {
    await compare(`<div id="owner" name="owner" style="display:flex;width:287.5px;height:187.5px;
      align-items:center;justify-content:space-evenly;background:#eee">
      <div id="authored" name="authored" style="display:flex;flex-direction:column;
        align-items:center;justify-content:center;width:137.5px;height:71.5px;background:#cef">
        Centered text</div></div>`, [{ scale: 1 }, { scale: 1.5 }, { scale: 2 }]);
  });
});
