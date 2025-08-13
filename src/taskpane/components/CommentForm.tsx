import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { MentionsInput, Mention } from "react-mentions";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";

// SharePoint Config
const siteId = "8314c8ba-c25a-4a02-bf25-d6238949ac8f";
const listId = "5f59364d-9808-4d26-8e04-2527b4fc403e";
const siteUrl = "https://jwelectricalsupply.sharepoint.com/sites/allcompany";
const tenantHost = "jwelectricalsupply.sharepoint.com";

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
  const fileInputRef = useRef<HTMLInputElement>(null);
  const commentsEndRef = useRef<HTMLDivElement>(null);

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
    if (Office.context.platform === Office.PlatformType.PC) {
      OfficeRuntime.storage.setItem("desktopMode", "true");
      document.cookie = `SameSite=None; Secure; domain=${window.location.hostname}`;
    }
  }, []);

  useEffect(() => {
    Office.onReady(async (info) => {
      if (info.host === Office.HostType.Outlook) {
        await waitForMailboxItem();
        setLoading(true);
        const item = Office.context.mailbox.item;

        item.body.getAsync(Office.CoercionType.Html, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            let bodyContent = result.value as string;
            if (Office.context.platform === Office.PlatformType.PC) {
              bodyContent = bodyContent.replace(/<meta\s[^>]*>/gi, "");
            }
            const match = bodyContent.match(/CONVERSATION_ID:([a-zA-Z0-9\-]+)/);
            let convId = match?.[1] || (item as any).conversationId;
            if (convId) {
              setConversationId(convId);
              await fetchCommentsFromSharePoint(convId);
            }
          }
          setLoading(false);
        });
      }
    });
  }, []);

  // Auto-refresh comments every 15 seconds
  useEffect(() => {
    if (!conversationId) return () => {};
    const interval = setInterval(() => {
      fetchCommentsFromSharePoint(conversationId);
    }, 15000);
    return () => clearInterval(interval);
  }, [conversationId]);

  useEffect(() => {
    fetchUsers();
  }, []);

  const initializeMsal = async () => {
    try {
      await msalInstance.initialize();
    } catch {}
  };

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

  const getSharePointToken = async (): Promise<string> => {
    await initializeMsal();
    let accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      const loginResp = await msalInstance.loginPopup({
        scopes: ["User.Read", "Mail.Send", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      });
      accounts = [loginResp.account];
    }
    const spScope =
      Office.context.platform === Office.PlatformType.PC
        ? `https://${tenantHost}/AllSites.FullControl`
        : `https://${tenantHost}/.default`;
    try {
      const resp = await msalInstance.acquireTokenSilent({
        scopes: [spScope],
        account: accounts[0],
      });
      return resp.accessToken;
    } catch {
      const resp = await msalInstance.acquireTokenPopup({
        scopes: [spScope],
      });
      return resp.accessToken;
    }
  };

  const fetchUsers = async () => {
    try {
      const token = await getGraphToken();
      const response = await fetch("https://graph.microsoft.com/v1.0/users?$top=50", {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) throw new Error("Failed to fetch users");
      const data = await response.json();
      setPeople(
        data.value.map((user: any) => ({
          id: user.mail || user.userPrincipalName,
          display: user.displayName,
        }))
      );
    } catch {}
  };

  const fetchCommentsFromSharePoint = async (convId?: string): Promise<void> => {
    const id = convId || conversationId;
    if (!id) return;
    setLoading(true);
    try {
      const graphToken = await getGraphToken();
      const spToken = await getSharePointToken();
      const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$orderby=createdDateTime asc`;
      const response = await fetch(url, { headers: { Authorization: `Bearer ${graphToken}` } });
      const data = await response.json();
      const comments = await Promise.all(
        data.value
          .filter((item: any) => item.fields?.EmailID === id)
          .map(async (item: any) => {
            const fields: ICommentFields = item.fields;
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
              fields.Attachments = attJson.d?.results || [];
            } else {
              fields.Attachments = [];
            }
            return fields;
          })
      );
      setCommentHistory(comments);
    } catch {}
    setLoading(false);
    return;
  };

  const stripMentionsFromComment = (input: string): string => {
    return input.replace(/@\[[^\]]+\]\([^)]+\)/g, "").trim();
  };

  const handleSaveAndShare = async () => {
    if (!comment.trim()) {
      alert("Please add a comment before saving.");
      return;
    }
    setSending(true);
    try {
      // save comment code here (unchanged from your version)
      setComment("");
      setSelectedFiles([]);
      if (fileInputRef.current) fileInputRef.current.value = "";
      await fetchCommentsFromSharePoint();
    } finally {
      setSending(false);
    }
  };

  useEffect(() => {
    if (!loading && commentsEndRef.current) {
      commentsEndRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [commentHistory, loading]);

  return (
    <div style={{ padding: "1rem", fontFamily: "Segoe UI, sans-serif", fontSize: "14px" }}>
      {/* Comments Section */}
      <div
        style={{
          marginBottom: "1rem",
          background: "#f9f9f9",
          padding: "10px",
          borderRadius: "6px",
          maxHeight: "calc(100vh - 180px)",
          overflowY: "auto",
          position: "relative",
          paddingBottom: "80px",
        }}
      >
        {loading ? (
          <Spinner size={SpinnerSize.medium} label="Loading comments..." />
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
              <div
                style={{
                  width: 40,
                  height: 40,
                  borderRadius: "50%",
                  backgroundColor: "#dfe1e5",
                  textAlign: "center",
                  lineHeight: "40px",
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
                {c.Attachments && c.Attachments.length > 0 && (
                  <div style={{ marginTop: 8 }}>
                    <strong style={{ fontSize: 12 }}>Attachments:</strong>
                    <ul style={{ margin: "6px 0 0 0", paddingLeft: 18 }}>
                      {c.Attachments.map((file: any, idx: number) => (
                        <li key={idx}>
                          <a
                            href={`${siteUrl}${file.ServerRelativeUrl}`}
                            target="_blank"
                            rel="noopener noreferrer"
                          >
                            {file.FileName}
                          </a>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            </div>
          ))
        ) : (
          <div style={{ color: "#888" }}>No comments yet.</div>
        )}
        <div ref={commentsEndRef} />
      </div>

      {/* Fixed Input */}
      <div
        style={{
          position: "fixed",
          bottom: 0,
          left: 0,
          right: 0,
          background: "#fff",
          padding: "10px",
          borderTop: "1px solid #ccc",
          display: "flex",
          flexDirection: "column",
          gap: "8px",
        }}
      >
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

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <input
            type="file"
            multiple
            ref={fileInputRef}
            onChange={(e) => setSelectedFiles(Array.from(e.target.files || []))}
          />
          <button
            onClick={handleSaveAndShare}
            disabled={sending}
            style={{
              backgroundColor: sending ? "#a0a0a0" : "#0078D4",
              color: "#fff",
              border: "none",
              padding: "10px 16px",
              borderRadius: "5px",
              cursor: sending ? "not-allowed" : "pointer",
            }}
          >
            {sending ? "Sending..." : "Send"}
          </button>
        </div>
      </div>
    </div>
  );
};

export default CommentForm;
