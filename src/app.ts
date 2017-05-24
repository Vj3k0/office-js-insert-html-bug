/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#run').click(run);
    });
  };

  async function run() {
    
    await Word.run(async (context) => {

      let bibliography = `
        <div class="csl-bib-body" style="line-height: 2; ">
          <div class="csl-entry" style="clear: left; ">
            <div class="csl-left-margin" style="float: left; padding-right: 0.5em; text-align: right; width: 1em;">1.</div><div class="csl-right-inline" style="margin: 0 .4em 0 1.5em;">Hisakata, R., Nishida, S. ’ya &amp; Johnston, A. An adaptable metric shapes perceptual space. <i>Curr. Biol.</i> <b>26,</b> 1911–1915 (2016).</div>
          </div>
          <div class="csl-entry" style="clear: left; ">
            <div class="csl-left-margin" style="float: left; padding-right: 0.5em; text-align: right; width: 1em;">2.</div><div class="csl-right-inline" style="margin: 0 .4em 0 1.5em;">Musk, E. The secret Tesla Motors master plan (just between you and me). <i>Tesla Blog</i> (2006). Available at: https://www.tesla.com/blog/secret-tesla-motors-master-plan-just-between-you-and-me. (Accessed: 29th September 2016)</div>
          </div>
          <div class="csl-entry" style="clear: left; ">
            <div class="csl-left-margin" style="float: left; padding-right: 0.5em; text-align: right; width: 1em;">3.</div><div class="csl-right-inline" style="margin: 0 .4em 0 1.5em;">Hogue, C. W. V. in <i>Bioinformatics</i> (eds. Baxevanis, A. D. &amp; Ouellette, B. F. F.) 83–109 (Wiley-Interscience, 2001).</div>
          </div>
          <div class="csl-entry" style="clear: left; ">
            <div class="csl-left-margin" style="float: left; padding-right: 0.5em; text-align: right; width: 1em;">4.</div><div class="csl-right-inline" style="margin: 0 .4em 0 1.5em;">Sambrook, J. &amp; Russell, D. W. <i>Molecular cloning: a laboratory manual</i>. (CSHL Press, 2001).</div>
          </div>
        </div>
      `;

      let range = context.document.getSelection();

      // Create new content control and define content and type
      let cc = range.insertContentControl();
      let ccRange = cc.insertHtml(bibliography, 'replace');

      /**
       * Insert your Word code here
       */
      await context.sync();
    });
    
  }
})();
