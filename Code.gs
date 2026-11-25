/**
 * Femboy Syndicate DKP - public web page.
 * Bound to the "Femboy Syndicate DKP and Vault HQ" sheet.
 * Read-only: nothing here writes to the sheet.
 * Which sections are public is controlled on the Config tab (K5:L9).
 */

function doGet(e) {
  // ?data=1 returns JSON, used by the GitHub Pages version of the site.
  if (e && e.parameter && e.parameter.data) {
    return ContentService.createTextOutput(JSON.stringify(getPublicData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Femboy Syndicate DKP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function settings_(ss) {
  var rows = ss.getSheetByName('Config').getRange('K5:L9').getValues();
  var on = {};
  rows.forEach(function (r) {
    on[String(r[0]).trim()] = String(r[1]).trim().toLowerCase() === 'yes';
  });
  return on;
}

function num_(v) {
  return typeof v === 'number' ? v : null;
}

/** Called by the page. Only returns sections switched on in Config. */
function getPublicData() {
  var ss = SpreadsheetApp.getActive();
  var on = settings_(ss);
  var cores = ss.getSheetByName('Config').getRange('A5:A12').getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (c) { return c !== ''; });

  var out = { guild: 'Femboy Syndicate', updated: new Date().toISOString(), on: on, cores: cores };

  if (on['Standings']) {
    out.standings = cores.map(function (core) {
      var sh = ss.getSheetByName(core);
      var rows = [];
      if (sh && sh.getLastRow() >= 3) {
        rows = sh.getRange(3, 1, sh.getLastRow() - 2, 5).getValues()
          .filter(function (r) { return r[0] !== '' && typeof r[4] === 'number'; })
          .map(function (r) {
            return { name: String(r[0]), earned: r[1], spent: r[3], balance: r[4] };
          });
      }
      return { core: core, rows: rows };
    });
  }

  if (on['Raid comp']) {
    var comp = ss.getSheetByName('Comp');
    var roles = comp.getRange('B7:L9').getValues().map(function (r) {
      return {
        core: String(r[0]), raiders: r[1],
        tanks: r[2], tankTarget: r[3],
        healers: r[4], healerTarget: r[5],
        dps: r[6], dpsTarget: r[7],
        melee: r[8], ranged: r[9], need: String(r[10])
      };
    });
    var classHeader = comp.getRange('C27:E27').getValues()[0].map(String);
    var classes = comp.getRange('B28:F36').getValues().map(function (r) {
      return { cls: String(r[0]), counts: [r[1], r[2], r[3]], total: r[4] };
    });
    out.comp = { roles: roles, classCores: classHeader, classes: classes };
  }

  if (on['Vault']) {
    var v = ss.getSheetByName('Vault_Items');
    var last = v.getLastRow();
    var items = [];
    if (last >= 2) {
      items = v.getRange(2, 1, last - 1, 7).getValues()
        .filter(function (r) { return r[0] !== '' && r[3] > 0 && String(r[6]) === 'For sale'; })
        .map(function (r) {
          return { item: String(r[0]), left: r[3], dkp: r[4] === '' ? null : r[4], gold: r[5] === '' ? null : r[5] };
        });
    }
    out.vault = items;
  }

  var dd = ss.getSheetByName('Dashboard_Data');
  if (on['Top donors']) {
    out.donors = dd.getRange('J2:K9').getValues()
      .filter(function (r) { return r[0] !== '' && typeof r[1] === 'number' && r[1] > 0; })
      .map(function (r) { return { name: String(r[0]), gold: r[1] }; });
  }
  if (on['Guild bank']) {
    out.bank = {
      gold: num_(dd.getRange('T6').getValue()),
      donated: num_(dd.getRange('T2').getValue()),
      units: num_(dd.getRange('T8').getValue())
    };
  }
  return out;
}

/**
 * Optional: post the top 10 for each core to Discord.
 * 1. Project Settings > Script Properties > add DISCORD_WEBHOOK = your webhook URL.
 * 2. Run this by hand, or add a time trigger (Triggers > Add trigger).
 * The URL never goes in a cell. Mentions are blocked.
 */
function postStandingsToDiscord() {
  var url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK');
  if (!url) throw new Error('Add DISCORD_WEBHOOK in Project Settings > Script Properties.');
  var d = getPublicData();
  if (!d.standings) throw new Error('Standings are switched off on the Config tab.');
  var colors = [0x257de0, 0xf47230, 0x00b894];
  var embeds = d.standings.slice(0, 3).map(function (s, i) {
    var top = s.rows.slice().sort(function (a, b) { return b.balance - a.balance; }).slice(0, 10);
    return {
      title: s.core + ' DKP - top 10',
      color: colors[i % 3],
      description: top.length
        ? top.map(function (r, j) { return (j + 1) + '. ' + r.name + ' - ' + r.balance; }).join('\n')
        : 'No players yet',
      timestamp: new Date().toISOString()
    };
  });
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      username: 'Femboy Syndicate DKP',
      allowed_mentions: { parse: [] },
      embeds: embeds
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Discord returned ' + res.getResponseCode());
  }
}
