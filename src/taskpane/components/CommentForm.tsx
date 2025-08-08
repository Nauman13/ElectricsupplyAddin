import * as React from "react";
import { useState, useEffect } from "react";
import { MentionsInput, Mention } from "react-mentions";
import { PublicClientApplication, AuthenticationResult } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";

// SharePoint Config
const siteId = "8314c8ba-c25a-4a02-bf25-d6238949ac8f";
const listId = "5f59364d-9808-4d26-8e04-2527b4fc403e";
const siteUrl = "https://jwelectricalsupply.sharepoint.com/sites/allcompany";
const tenantHost = "jwelectricalsupply.sharepoint.com"; // used for SP token scope

const msalInstance = new PublicClientApplication(msalConfig);

interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

interface ICommentFields {
  Title: string;
  EmailID: string;
  Comment: string;
  MentionedUsers?: string;
  CreatedBy?: string;
  CreatedDate?: string;
  Attachments?: IAttachment[];
}

const CommentForm: React.FC = () => {
  const [comment, setComment] = useState<string>("");
  const [commentHistory, setCommentHistory] = useState<ICommentFields[]>([]);
  const [people, setPeople] = useState<any[]>([]);
  const [mentionedEmails, setMentionedEmails] = useState<string[]>([]);
  const [conversationId, setConversationId] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [sending, setSending] = useState<boolean>(false);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  // Wait until Office item is available
  const waitForMailboxItem = (): Promise<void> => {
    return new Promise((resolve) => {
      const check = () => {
        if (Office.context?.mailbox?.item) resolve();
        else setTimeout(check, 100);
      };
      check();
    });
  };

  useEffect(() => {
    Office.onReady(async (info) => {
      if (info.host === Office.HostType.Outlook) {
        await waitForMailboxItem();
        setLoading(true);
        const item = Office.context.mailbox.item;

        item.body.getAsync(Office.CoercionType.Html, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const bodyContent = result.value as string;
            const match = bodyContent.match(/CONVERSATION_ID:([a-zA-Z0-9\-]+)/);
            const convId = match?.[1] || (item as any).conversationId;
            if (convId) {
              setConversationId(convId);
              // fetch comments + attachments
              await fetchCommentsFromSharePoint(convId);
            }
          } else {
            console.error("Failed to get body:", result.error);
          }
          setLoading(false);
        });
      }
    });
  }, []);

  useEffect(() => {
    fetchUsers();
  }, []);

  // -------------------------
  // Auth helpers (MSAL)
  // -------------------------
  const initializeMsal = async () => {
    try {
      await msalInstance.initialize();
    } catch (e) {
      console.warn("msal initialize warning:", e);
    }
  };

  // Get a Graph token (for Graph calls like create item & send mail)
  const getGraphToken = async (): Promise<string> => {
    await initializeMsal();
    let accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      const loginResp = await msalInstance.loginPopup({
        scopes: ["User.Read", "Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      });
      accounts = [loginResp.account];
    }

    const resp = await msalInstance.acquireTokenSilent({
      scopes: ["User.Read", "Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      account: accounts[0],
    });

    return resp.accessToken;
  };

  // Get a SharePoint resource token (audience = https://{tenant}.sharepoint.com)
  const getSharePointToken = async (): Promise<string> => {
    await initializeMsal();
    let accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      const loginResp = await msalInstance.loginPopup({
        scopes: ["User.Read", "Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      });
      accounts = [loginResp.account];
    }

    const spScope = `https://${tenantHost}/.default`;

    try {
      const resp = await msalInstance.acquireTokenSilent({
        scopes: [spScope],
        account: accounts[0],
      });
      return resp.accessToken;
    } catch (err) {
      const resp = await msalInstance.acquireTokenPopup({
        scopes: [spScope],
      });
      return resp.accessToken;
    }
  };

  // -------------------------
  // Email forward (Graph)
  // -------------------------
  const forwardOriginalEmailToMentionedUsers = async () => {
    if (mentionedEmails.length === 0) return;

    const token = await getGraphToken();

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get email body:", result.error);
        return;
      }

      let originalBody = result.value as string;
      if (!originalBody.includes("CONVERSATION_ID:")) {
        originalBody += `<p style="color:#fff;font-size:1px">CONVERSATION_ID:${conversationId}</p>`;
      }

      const subject = Office.context.mailbox.item.subject;
      const from = Office.context.mailbox.item.from?.emailAddress || "noreply@domain.com";

      const toRecipients = mentionedEmails.map((email) => ({ emailAddress: { address: email } }));

      const emailPayload = {
        message: {
          subject: ` ${subject}`,
          body: {
            contentType: "HTML",
            content: `
              <p>Hello,</p>
              <p>You were mentioned in a conversation. Here's the original email:</p>
              <hr />
              <p><strong>From:</strong> ${from}</p>
              <p><strong>Subject:</strong> ${subject}</p>
              <hr />
              ${originalBody}
              <hr />
              <p>You can open the Outlook Add-in to view and add comments.</p>
            `,
          },
          toRecipients,
        },
        saveToSentItems: true,
      };

      const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(emailPayload),
      });

      if (!res.ok) {
        const errText = await res.text();
        console.error("Failed to send mail:", errText);
      } else {
        console.log("Successfully sent forward-style email to mentioned users.");
      }
    });
  };

  // -------------------------
  // Users list for Mention suggestions
  // -------------------------
  const fetchUsers = async () => {
    try {
      const token = await getGraphToken();
      const response = await fetch("https://graph.microsoft.com/v1.0/users?$top=50", {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) throw new Error("Failed to fetch users");

      const data = await response.json();
      const usersData = data.value.map((user: any) => ({
        id: user.mail || user.userPrincipalName,
        display: user.displayName,
        email: user.mail || user.userPrincipalName,
      }));

      setPeople(usersData);
    } catch (error) {
      console.error("Error fetching users:", error);
    }
  };

  // -------------------------
  // Fetch comments (Graph) + fetch attachments per item (SharePoint REST + SP token)
  // -------------------------
  const fetchCommentsFromSharePoint = async (convId?: string) => {
    const id = convId || conversationId;
    if (!id) return;

    setLoading(true);
    try {
      const graphToken = await getGraphToken();
      const spToken = await getSharePointToken();

      const emailIdValue = encodeURIComponent(id);
      const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$filter=fields/EmailID eq '${emailIdValue}'&$orderby=createdDateTime asc`;

      const response = await fetch(url, {
        headers: { Authorization: `Bearer ${graphToken}` },
      });

      if (!response.ok) {
        const txt = await response.text();
        throw new Error(`Failed to fetch list items: ${txt}`);
      }

      const data = await response.json();

      const comments = await Promise.all(
        data.value.map(async (item: any) => {
          const fields: ICommentFields = item.fields || ({} as ICommentFields);

          // Fetch attachments for this list item using SharePoint REST and spToken
          try {
            const attRes = await fetch(
              `${siteUrl}/_api/web/lists(guid'${listId}')/items(${item.id})/AttachmentFiles`,
              {
                headers: {
                  Authorization: `Bearer ${spToken}`,
                  Accept: "application/json;odata=verbose",
                },
              }
            );

            if (attRes.ok) {
              const attJson = await attRes.json();
              // odata verbose shape: attJson.d.results
              fields.Attachments = attJson.d?.results || [];
            } else {
              console.warn(
                `Attachment fetch failed for item ${item.id}:`,
                attRes.status,
                await attRes.text()
              );
              fields.Attachments = [];
            }
          } catch (attErr) {
            console.error("Attachment fetch error:", attErr);
            fields.Attachments = [];
          }

          return fields;
        })
      );

      setCommentHistory(comments);
    } catch (error) {
      console.error("Error fetching comments:", error);
    } finally {
      setLoading(false);
    }
  };

  // -------------------------
  // Mention helpers
  // -------------------------
  const stripMentionsFromComment = (input: string): string => {
    return input.replace(/@\[[^\]]+\]\([^)]+\)/g, "").trim();
  };

  const extractMentionData = (input: string) => {
    const mentionRegex = /@\[([^\]]+)\]\(([^)]+)\)/g;
    const displayNames: string[] = [];
    const emails: string[] = [];

    let match;
    while ((match = mentionRegex.exec(input)) !== null) {
      displayNames.push(match[1]);
      emails.push(match[2]);
    }

    return { displayNames, emails };
  };

  // -------------------------
  // Save comment -> create list item (Graph) then upload attachments (SP REST using spToken)
  // -------------------------
  const saveCommentToSharePoint = async () => {
    const graphToken = await getGraphToken();
    const spToken = await getSharePointToken();

    const plainComment = stripMentionsFromComment(comment);
    const { displayNames, emails } = extractMentionData(comment);

    const fieldsData: any = {
      Title: "Email Comment",
      EmailID: conversationId,
      Comment: plainComment,
      MentionedUsers: displayNames.join(", "),
      CreatedBy: Office.context.mailbox.userProfile.displayName,
      CreatedDate: new Date().toISOString(),
    };

    // Create List Item via Graph
    const createItemRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${graphToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ fields: fieldsData }),
      }
    );

    if (!createItemRes.ok) {
      const errText = await createItemRes.text();
      throw new Error(`Failed creating list item: ${errText}`);
    }

    const itemData = await createItemRes.json();
    const itemId = itemData.id;

    // Upload attachments to this item using SharePoint REST (SP token)
    for (const file of selectedFiles) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const uploadUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(
          file.name
        )}')`;

        const uploadRes = await fetch(uploadUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${spToken}`,
            Accept: "application/json;odata=verbose",
            // Do not set Content-Type for binary body in many browsers; let fetch handle it.
          },
          body: arrayBuffer,
        });

        if (!uploadRes.ok) {
          const t = await uploadRes.text();
          console.warn(`Upload failed for ${file.name}:`, uploadRes.status, t);
        }
      } catch (e) {
        console.error("Upload attachment error:", e);
      }
    }

    setMentionedEmails(emails);
  };

  // -------------------------
  // Save + refresh + optionally notify mentioned users
  // -------------------------
  const handleSaveAndShare = async () => {
    if (!comment.trim()) {
      alert("Please add a comment before saving.");
      return;
    }

    setSending(true);
    try {
      await saveCommentToSharePoint();
      await fetchCommentsFromSharePoint();

      if (mentionedEmails.length > 0) {
        await forwardOriginalEmailToMentionedUsers();
      }
      setComment("");
      setMentionedEmails([]);
      setSelectedFiles([]);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (error) {
      console.error("Error saving comment:", error);
      alert("Error saving comment. Check console for details.");
    } finally {
      setSending(false);
    }
  };
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  return (
    <div style={{ padding: "1rem", fontFamily: "Segoe UI, sans-serif", fontSize: "14px" }}>
      <div
        style={{
          marginBottom: "1rem",
          background: "#f9f9f9",
          padding: "10px",
          borderRadius: "6px",
          maxHeight: "450px",
          overflowY: "auto",
          position: "relative",
          minHeight: "100px",
        }}
      >
        {loading ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              height: "100px",
            }}
          >
            <Spinner size={SpinnerSize.medium} label="Loading comments..." />
          </div>
        ) : commentHistory.length > 0 ? (
          commentHistory.map((c, index) => (
            <div
              key={index}
              style={{
                display: "flex",
                alignItems: "flex-start",
                marginBottom: "15px",
                padding: "10px",
                borderRadius: "8px",
                backgroundColor: "#f4f6f9",
              }}
            >
              {/* Initial */}
              <div
                style={{
                  width: 40,
                  height: 40,
                  borderRadius: "50%",
                  backgroundColor: "#dfe1e5",
                  textAlign: "center",
                  lineHeight: "30px",
                  fontWeight: "bold",
                  fontSize: "16px",
                  marginRight: "10px",
                }}
              >
                {c.CreatedBy?.charAt(0).toUpperCase()}
              </div>

              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 600 }}>{c.CreatedBy}</div>
                <div style={{ margin: "5px 0", whiteSpace: "pre-wrap" }}>{c.Comment}</div>
                {c.MentionedUsers && (
                  <div style={{ fontSize: "12px" }}>
                    {c.MentionedUsers.split(",").map((name: string, i: number) => (
                      <span
                        key={i}
                        style={{
                          backgroundColor: "#e6f0ff",
                          color: "#1a73e8",
                          padding: "2px 6px",
                          borderRadius: "4px",
                          marginRight: "5px",
                        }}
                      >
                        @{name.trim()}
                      </span>
                    ))}
                  </div>
                )}
                <div style={{ fontSize: "11px", color: "#888", marginTop: "4px" }}>
                  {new Date(c.CreatedDate).toLocaleString()}
                </div>

                {/* ---------- ATTACHMENTS: SHOW LINKS ONLY ---------- */}
                {c.Attachments && c.Attachments.length > 0 && (
                  <div style={{ marginTop: 8 }}>
                    <strong style={{ fontSize: 12 }}>Attachments:</strong>
                    <ul style={{ margin: "6px 0 0 0", paddingLeft: 18 }}>
                      {c.Attachments.map((file: any, idx: number) => {
                        // Build full URL from ServerRelativeUrl
                        const fileUrl = `${siteUrl}${file.ServerRelativeUrl}`;
                        return (
                          <li key={idx} style={{ marginBottom: 6 }}>
                            <a
                              href={fileUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              style={{ color: "#0078d4", textDecoration: "underline" }}
                            >
                              {file.FileName}
                            </a>
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                )}
              </div>
            </div>
          ))
        ) : (
          <div style={{ color: "#888" }}>No comments yet.</div>
        )}
      </div>

      <div style={{ marginBottom: "10px", borderRadius: "10px" }}>
        <MentionsInput
          value={comment}
          onChange={(e) => setComment(e.target.value)}
          placeholder="Add internal comment..."
          style={{
            control: {
              backgroundColor: "#fff",
              fontSize: 14,
              padding: "8px",
              borderRadius: "10px",
              border: "1px solid #ccc",
              minHeight: "40px",
              maxHeight: "60px",
              overflowY: "auto",
            },
            input: {
              margin: 0,
              paddingLeft: "10px",
              borderRadius: "10px",
              outline: "none",
              border: "none",
            },
            highlighter: {
              overflow: "hidden",
            },
            suggestions: {
              list: {
                backgroundColor: "#fff",
                border: "1px solid #ccc",
                fontSize: 14,
              },
              item: {
                padding: "5px 10px",
                borderBottom: "1px solid #eee",
                "&focused": {
                  backgroundColor: "#e6f0ff",
                },
              },
            },
          }}
        >
          <Mention
            trigger="@"
            data={people}
            displayTransform={(_id, display) => `@${display}`}
            appendSpaceOnAdd
            onAdd={(id: string) => {
              if (!mentionedEmails.includes(id)) {
                setMentionedEmails([...mentionedEmails, id]);
              }
            }}
          />
        </MentionsInput>
        <input
          type="file"
          multiple
          ref={fileInputRef}
          onChange={(e) => setSelectedFiles(Array.from(e.target.files || []))}
        />
      </div>

      <button
        onClick={handleSaveAndShare}
        disabled={sending}
        style={{
          backgroundColor: sending ? "#a0a0a0" : "#0078D4",
          color: "#fff",
          border: "none",
          padding: "10px 16px",
          borderRadius: "5px",
          float: "right",
          display: "flex",
          alignItems: "center",
          gap: "8px",
          cursor: sending ? "not-allowed" : "pointer",
        }}
      >
        {sending ? (
          <>
            <Spinner size={SpinnerSize.xSmall} />
            Sending...
          </>
        ) : (
          "Send"
        )}
      </button>
    </div>
  );
};

export default CommentForm;
