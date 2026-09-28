Office.onReady(function () {});

function escapeHtml(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function guessCompanyFromBody(bodyText) {
  if (!bodyText) return null;
  var lines = bodyText.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var t = lines[i].trim();
    if (!t) continue;
    if (t.indexOf("弘進ゴム") !== -1) continue;
    if (/(株式会社|有限会社|合同会社|\(株\)|㈱)/.test(t) && t.length <= 40) {
      return t;
    }
  }
  return null;
}

function insertGreeting(event) {
  var item = Office.context.mailbox.item;

  function buildAndInsert(name, company) {
    var html = "";
    if (company) {
      html += "<div>" + escapeHtml(company) + "</div>";
    }
    html +=
      "<div>" + escapeHtml(name) + "様</div>" +
      "<div>&nbsp;</div>" +
      "<div>いつもお世話になっております。</div>" +
      "<div>&nbsp;</div>";
    item.body.prependAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () {
        event.completed();
      }
    );
  }

  item.to.getAsync(function (toResult) {
    var name = "";
    if (
      toResult.status === Office.AsyncResultStatus.Succeeded &&
      toResult.value &&
      toResult.value.length > 0
    ) {
      var r0 = toResult.value[0];
      name = (r0.displayName || "").trim();
      if (!name || name.indexOf("@") !== -1) {
        var email = r0.emailAddress || "";
        name = email.split("@")[0];
      }
    }
    if (!name) name = "ご担当者";

    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      var bodyText =
        bodyResult.status === Office.AsyncResultStatus.Succeeded
          ? bodyResult.value
          : "";
      var company = guessCompanyFromBody(bodyText);
      buildAndInsert(name, company);
    });
  });
}

Office.actions.associate("insertGreeting", insertGreeting);
