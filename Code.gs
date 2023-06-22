function getOAuthService() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const clientSecret = scriptProperties.getProperty(
    "TRACKER_API_CLIENT_SECRET"
  );
  const clientId = scriptProperties.getProperty("TRACKER_API_CLIENT_ID");
  const baseUrl = scriptProperties.getProperty("TRACKER_API_BASE_URL");

  return (
    OAuth2.createService("getomniscience")
      .setAuthorizationBaseUrl(baseUrl + "/o/oauth2/auth")
      .setTokenUrl(baseUrl + "/o/oauth2/token")
      .setClientId(clientId)
      .setClientSecret(clientSecret)
      // .setScope('SERVICE_SCOPE_REQUESTS')
      .setCallbackFunction("authCallback")
      .setCache(CacheService.getUserCache())
      .setPropertyStore(PropertiesService.getUserProperties())
      .setLock(LockService.getUserLock())
      //getActiveUser?
      .setParam("login_hint", Session.getEffectiveUser().getEmail())
  );
}

function create3PAuthorizationUi() {
  var service = getOAuthService();
  var authUrl = service.getAuthorizationUrl();

  const scriptProperties = PropertiesService.getScriptProperties();
  const baseUrl = scriptProperties.getProperty("TRACKER_API_BASE_URL");

  var logoutButton = CardService.newTextButton()
    .setText("Logout")
    .setOpenLink(
      CardService.newOpenLink()
        .setUrl(baseUrl + "/logout")
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON)
    );

  var authButton = CardService.newTextButton()
    .setText("Begin Auth")
    .setAuthorizationAction(
      CardService.newAuthorizationAction().setAuthorizationUrl(authUrl)
    );

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Authorization Required"))
    .addSection(
      CardService.newCardSection()
        .setHeader("This add-on needs access to your Email Tracker account.")
        .addWidget(
          CardService.newButtonSet()
            .addButton(authButton)
            .addButton(logoutButton)
        )
    )
    .build();
  return [card];
}

/**
 * Attempts to access a non-Google API using a constructed service
 * object.
 *
 * If your add-on needs access to non-Google APIs that require OAuth,
 * you need to implement this method. You can use the OAuth1 and
 * OAuth2 Apps Script libraries to help implement it.
 *
 * @param {String} url         The URL to access.
 * @param {String} method_opt  The HTTP method. Defaults to GET.
 * @param {Object} headers_opt The HTTP headers. Defaults to an empty
 *                             object. The Authorization field is added
 *                             to the headers in this method.
 * @return {HttpResponse} the result from the UrlFetchApp.fetch() call.
 */
function accessProtectedResource(url, method_opt, headers_opt, json) {
  var service = getOAuthService();
  var maybeAuthorized = service.hasAccess();
  if (maybeAuthorized) {
    // A token is present, but it may be expired or invalid. Make a
    // request and check the response code to be sure.

    // Make the UrlFetch request and return the result.
    var accessToken = service.getAccessToken();
    var method = method_opt || "get";
    var headers = headers_opt || {};
    headers["Authorization"] = Utilities.formatString("Bearer %s", accessToken);

    var fetchOptions = {
      headers: headers,
      method: method,
      muteHttpExceptions: true, // Prevents thrown HTTP exceptions.
    };

    if (json) {
      fetchOptions["contentType"] = "application/json";
      fetchOptions["payload"] = JSON.stringify(json);
    }

    var resp = UrlFetchApp.fetch(url, fetchOptions);

    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return resp.getContentText("utf-8"); // Success
    } else if (code == 401 || code == 403) {
      // Not fully authorized for this action.
      maybeAuthorized = false;
    } else {
      // Handle other response codes by logging them and throwing an
      // exception.
      console.error(
        "Backend server error (%s): %s",
        code.toString(),
        resp.getContentText("utf-8")
      );
      throw "Backend server error: " + code;
    }
  }

  if (!maybeAuthorized) {
    // Invoke the authorization flow using the default authorization
    // prompt card.
    CardService.newAuthorizationException()
      .setAuthorizationUrl(service.getAuthorizationUrl())
      .setResourceDisplayName("Email Tracker Login")
      .setCustomUiCallback("create3PAuthorizationUi")
      .throwException();
  }
}

function onGmailMessageOpen(e) {
  console.log("onGmailMessageOpen", e);

  //var service = getOAuthService();
  //var userId = service.getStorage().getValue("userId");

  // Activate temporary Gmail scopes, in this case to allow
  // message metadata to be read.
  //var accessToken = e.gmail.accessToken;
  //GmailApp.setCurrentMessageAccessToken(accessToken);

  //var messageId = e.gmail.messageId;
  var threadId = e.gmail.threadId;
  //var message = GmailApp.getMessageById(messageId);
  //var subject = message.getSubject();
  //var sender = message.getFrom();

  //const emailId = message.getId();
  //const emailIdRemake = BigInt(`0x${emailId}`).toString(10);

  // Setting the access token with a gmail.addons.current.message.readonly
  // scope also allows read access to the other messages in the thread.
  // var thread = message.getThread();

  // Using this link can avoid the need to copy message or thread content
  // var threadLink = thread.getPermalink();

  const scriptProperties = PropertiesService.getScriptProperties();
  const baseUrl = scriptProperties.getProperty("TRACKER_API_BASE_URL");

  var threadResp = accessProtectedResource(
    `${baseUrl}/api/v1/threads/${threadId}/views/`
  );
  var threadViews = JSON.parse(threadResp).data;

  var composeAction =
    CardService.newAction().setFunctionName("createReplyDraft");
  var replyButton = CardService.newTextButton()
    .setText("Reply")
    .setComposeAction(
      composeAction,
      CardService.ComposedEmailType.REPLY_AS_DRAFT
    );

  var replyAllAction = CardService.newAction().setFunctionName(
    "createReplyAllDraft"
  );
  var replyButtonAll = CardService.newTextButton()
    .setText("Reply All")
    .setComposeAction(
      replyAllAction,
      CardService.ComposedEmailType.REPLY_AS_DRAFT
    );

  const formatView = (view) => {
    var ret = Utilities.formatDate(
      new Date(view.createdAt),
      e.userTimezone.id,
      "MM/dd/yyyy h:mm a"
    );

    const { clientIpGeo } = view;

    if (clientIpGeo != null) {
      if (clientIpGeo.data != null) {
        const geoData = clientIpGeo.data;
        const regionStr = geoData.region ?? geoData.regionCode;

        const extraLocaleTxt = geoData.isMobile === true ? '; mobile' : '';
        ret = `${ret} (${geoData.city}, ${regionStr}${extraLocaleTxt})`;
      }
      if (clientIpGeo.emailProvider != null) {
        ret = `${ret} (${clientIpGeo.emailProvider})`;
      }
    }

    return ret;
  };

  var promptText = (threadViews || [])
    .map((view) => formatView(view))
    .join("<br>");

  var fixedFooter = CardService.newFixedFooter()
    .setPrimaryButton(replyButton)
    .setSecondaryButton(replyButtonAll);

  // Create a card with a single card section and two widgets.
  // Be sure to execute build() to finalize the card construction.
  var exampleCard = CardService.newCardBuilder()
    // .setHeader(CardService.newCardHeader()
    //   .setTitle('Email Tracking'))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newKeyValue()
          .setTopLabel(
            `Views (${threadViews === undefined ? "n/a" : String(threadViews.length)
            })`
          )
          .setContent(promptText)
          .setMultiline(true)
      )
      // .addWidget(CardService.newButtonSet().addButton(replyButton))
      // .addWidget(CardService.newTextParagraph()
      //   .setText())
    )
    .setFixedFooter(fixedFooter)
    .build(); // Don't forget to build the Card!

  return [exampleCard];
}

/**
 *  Creates a draft email (with an attachment and inline image)
 *  as a reply to an existing message.
 *  @param {Object} e An event object passed by the action.
 *  @return {ComposeActionResponse}
 */
function createReplyDraft(e) {
  return createReplyXDraft(e, false);
}

function createReplyAllDraft(e) {
  return createReplyXDraft(e, true);
}

function createReplyXDraft(e, replyAll) {
  var service = getOAuthService();

  // Activate temporary Gmail scopes, in this case to allow
  // a reply to be drafted.
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  // Creates a draft reply.
  var messageId = e.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var emailSubject = message.getSubject();
  // console.log(e, subject);
  const scriptProperties = PropertiesService.getScriptProperties();
  const baseUrl = scriptProperties.getProperty("TRACKER_API_BASE_URL");
  const imageBaseUrl = `${baseUrl}/t`;
  var trackingSlug = service.getStorage().getValue("trackingSlug");
  var trackId = Utilities.getUuid();

  const url = `${imageBaseUrl}/${trackingSlug}/${trackId}/image.gif`;
  const trackingPixelHtml = `<div height="1" width="1" style="background-image: url('${url}');" data-src="${url}" class="tracker-img"></div>`;
  // console.log(trackingPixelHtml);

  var thread = message.getThread();
  var creator = replyAll ? thread.createDraftReplyAll : thread.createDraftReply;
  var draft = creator("", {
    htmlBody: trackingPixelHtml,
  });
  var threadId = e.gmail.threadId;
  const emailId = draft.getMessageId();
  // console.log(draft.getId(), draft.getMessageId())

  const selfLoadMitigation = false;

  const reportData = {
    emailId,
    threadId,
    trackId,
    emailSubject,
    selfLoadMitigation,
  };
  const reportUrl = `${baseUrl}/api/v1/trackers/`;
  accessProtectedResource(reportUrl, "post", null, reportData);

  // Return a built draft response. This causes Gmail to present a
  // compose window to the user, pre-filled with the content specified
  // above.
  return CardService.newComposeActionResponseBuilder()
    .setGmailDraft(draft)
    .build();
}

function doActionEmail(e) {
  console.log("doActionEmail", e);
  return [];
}

/**
 * Boilerplate code to determine if a request is authorized and returns
 * a corresponding HTML message. When the user completes the OAuth2 flow
 * on the service provider's website, this function is invoked from the
 * service. In order for authorization to succeed you must make sure that
 * the service knows how to call this function by setting the correct
 * redirect URL.
 *
 * The redirect URL to enter is:
 * https://script.google.com/macros/d/<Apps Script ID>/usercallback
 *
 * See the Apps Script OAuth2 Library documentation for more
 * information:
 *   https://github.com/googlesamples/apps-script-oauth2#1-create-the-oauth2-service
 *
 *  @param {Object} callbackRequest The request data received from the
 *                  callback function. Pass it to the service's
 *                  handleCallback() method to complete the
 *                  authorization process.
 *  @return {HtmlOutput} a success or denied HTML message to display to
 *          the user. Also sets a timer to close the window
 *          automatically.
 */
function authCallback(callbackRequest) {
  var service = getOAuthService();
  var authorized = service.handleCallback(callbackRequest);
  if (authorized) {
    var accessToken = service.getAccessToken();
    var headers = {};
    headers["Authorization"] = Utilities.formatString("Bearer %s", accessToken);
    var options = { method: "post", headers };
    const scriptProperties = PropertiesService.getScriptProperties();
    const baseUrl = scriptProperties.getProperty("TRACKER_API_BASE_URL");
    var resp = UrlFetchApp.fetch(baseUrl + "/api/v1/me", options);
    var data = JSON.parse(resp.getContentText());

    var storage = service.getStorage();
    storage.setValue("emailAccount", data.data.emailAccount);
    storage.setValue("trackingSlug", data.data.trackingSlug);
    storage.setValue("userId", data.data.id);

    return HtmlService.createHtmlOutput(
      "Success! <script>setTimeout(function() { top.window.close() }, 1);</script>"
    );
  } else {
    return HtmlService.createHtmlOutput("Denied");
  }
}

/**
 * Unauthorizes the non-Google service. This is useful for OAuth
 * development/testing.  Run this method (Run > resetOAuth in the script
 * editor) to reset OAuth to re-prompt the user for OAuth.
 */
function resetOAuth() {
  getOAuthService().reset();
}
