import * as React from "react";
import { useState, useEffect } from "react";
import { MentionsInput, Mention } from "react-mentions";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Icon } from "@fluentui/react/lib/Icon";

const siteId = "8314c8ba-c25a-4a02-bf25-d6238949ac8f";
const listId = "5f59364d-9808-4d26-8e04-2527b4fc403e";

const msalInstance = new PublicClientApplication(msalConfig);

const CommentForm: React.FC = () => {
  const [comment, setComment] = useState<string>("");
  const [commentHistory, setCommentHistory] = useState<any[]>([]);
  const [people, setPeople] = useState<any[]>([]);
  const [mentionedEmails, setMentionedEmails] = useState<string[]>([]);
  const [conversationId, setConversationId] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [sending, setSending] = useState<boolean>(false);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [fileInputKey, setFileInputKey] = useState<number>(0);

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
            const match = result.value.match(/CONVERSATION_ID:([a-zA-Z0-9\-]+)/);
            const convId = match?.[1] || item.conversationId;
            if (convId) {
              setConversationId(convId);
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

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setSelectedFiles([...selectedFiles, ...Array.from(e.target.files)]);
    }
  };

  const removeFile = (index: number) => {
    const newFiles = [...selectedFiles];
    newFiles.splice(index, 1);
    setSelectedFiles(newFiles);
  };

  const initializeMsal = async () => {
    await msalInstance.initialize();
  };

  const getAccessToken = async (): Promise<string> => {
    await initializeMsal();
    let accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
      const loginResponse = await msalInstance.loginPopup({
        scopes: ["User.Read", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      });
      accounts = [loginResponse.account];
    }

    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["User.Read", "Mail.ReadWrite", "Sites.ReadWrite.All"],
      account: accounts[0],
    });

    return tokenResponse.accessToken;
  };

  const fetchUsers = async () => {
    try {
      const token = await getAccessToken();
      const response = await fetch("https://graph.microsoft.com/v1.0/users?$top=50", {
        headers: { Authorization: `Bearer ${token}` },
      });

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

  const fetchCommentsFromSharePoint = async (convId?: string) => {
    const id = convId || conversationId;
    if (!id) return;

    setLoading(true);
    try {
      const token = await getAccessToken();
      const emailIdValue = encodeURIComponent(id);

      const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields,attachments&$filter=fields/EmailID eq '${emailIdValue}'&$orderby=createdDateTime asc`;

      const response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        console.error("Fetch comments failed:", await response.text());
        return;
      }

      const data = await response.json();
      const comments = data.value.map((item: any) => ({
        ...item.fields,
        Attachments: item.attachments || [],
        SharePointId: item.id,
      }));

      setCommentHistory(comments);
    } catch (error) {
      console.error("Error fetching comments:", error);
    } finally {
      setLoading(false);
    }
  };

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

  const saveCommentToSharePoint = async () => {
    const token = await getAccessToken();
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

    const itemRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ fields: fieldsData }),
      }
    );

    const item = await itemRes.json();

    // Upload attachments if any
    if (selectedFiles.length > 0) {
      for (const file of selectedFiles) {
        const arrayBuffer = await file.arrayBuffer();
        const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${item.id}/attachments/${encodeURIComponent(file.name)}/content`;

        await fetch(uploadUrl, {
          method: "PUT",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/octet-stream",
          },
          body: arrayBuffer,
        });
      }
    }

    setMentionedEmails(emails);
    return item.id; // Return the SharePoint item ID
  };

  const handleSaveAndShare = async () => {
    if (!comment.trim() && selectedFiles.length === 0) {
      alert("Please add a comment or attach a file before saving.");
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
      setFileInputKey((prev) => prev + 1); // Reset file input
    } catch (error) {
      console.error("Error saving comment:", error);
    } finally {
      setSending(false);
    }
  };

  const forwardOriginalEmailToMentionedUsers = async () => {
    if (mentionedEmails.length === 0) return;

    const token = await getAccessToken();

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get email body:", result.error);
        return;
      }

      let originalBody = result.value;

      if (!originalBody.includes("CONVERSATION_ID:")) {
        originalBody += `<p style="color:#fff;font-size:1px">CONVERSATION_ID:${conversationId}</p>`;
      }

      const subject = Office.context.mailbox.item.subject;
      const from = Office.context.mailbox.item.from?.emailAddress || "noreply@domain.com";

      const toRecipients = mentionedEmails.map((email) => ({
        emailAddress: { address: email },
      }));

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
        console.error("Failed to send mail:", await res.text());
      }
    });
  };

  const downloadAttachment = async (itemId: string, attachmentId: string, fileName: string) => {
    try {
      const token = await getAccessToken();
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/attachments/${attachmentId}/$value`,
        {
          headers: { Authorization: `Bearer ${token}` },
        }
      );

      if (!response.ok) {
        throw new Error("Failed to download attachment");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error downloading attachment:", error);
    }
  };

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
        }}
      >
        {loading ? (
          <div style={{ display: "flex", justifyContent: "center", height: "100px" }}>
            <Spinner size={SpinnerSize.medium} label="Loading comments..." />
          </div>
        ) : commentHistory.length > 0 ? (
          commentHistory.map((c, index) => (
            <div
              key={index}
              style={{
                marginBottom: "15px",
                padding: "10px",
                borderRadius: "8px",
                backgroundColor: "#f4f6f9",
              }}
            >
              <div style={{ fontWeight: 600 }}>{c.CreatedBy}</div>
              <div>{c.Comment}</div>

              {c.MentionedUsers && (
                <div style={{ fontSize: "12px", marginTop: "4px" }}>
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

              {c.Attachments && c.Attachments.length > 0 && (
                <div style={{ marginTop: "8px" }}>
                  {c.Attachments.map((att: any, idx: number) => (
                    <div
                      key={idx}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        marginBottom: "4px",
                        cursor: "pointer",
                      }}
                      onClick={() => downloadAttachment(c.SharePointId, att.id, att.name)}
                    >
                      <Icon
                        iconName="Attach"
                        style={{
                          fontSize: 14,
                          marginRight: "6px",
                          color: "#0078d4",
                        }}
                      />
                      <span style={{ fontSize: "12px", color: "#0078d4" }}>{att.name}</span>
                    </div>
                  ))}
                </div>
              )}

              <div style={{ fontSize: "11px", color: "#888", marginTop: "4px" }}>
                {new Date(c.CreatedDate).toLocaleString()}
              </div>
            </div>
          ))
        ) : (
          <div style={{ color: "#888" }}>No comments yet.</div>
        )}
      </div>

      <div style={{ position: "relative" }}>
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
              minHeight: "60px",
              paddingRight: "30px", // Make space for the paperclip icon
            },
            input: {
              margin: 0,
              paddingLeft: "10px",
              border: "none",
              outline: "none",
              minHeight: "40px",
            },
            suggestions: {
              list: { backgroundColor: "#fff", border: "1px solid #ccc", fontSize: 14 },
              item: { padding: "5px 10px" },
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

        {/* Paperclip icon positioned inside the textarea */}
        <label
          style={{
            position: "absolute",
            right: "10px",
            bottom: "10px",
            cursor: "pointer",
            zIndex: 1,
          }}
          title="Attach file"
        >
          <Icon iconName="Attach" style={{ fontSize: 16, color: "#666" }} />
          <input
            key={fileInputKey}
            type="file"
            multiple
            onChange={handleFileChange}
            style={{ display: "none" }}
          />
        </label>
      </div>

      {/* Selected files preview */}
      {selectedFiles.length > 0 && (
        <div
          style={{
            marginTop: "8px",
            padding: "8px",
            backgroundColor: "#f5f5f5",
            borderRadius: "4px",
          }}
        >
          {selectedFiles.map((file, i) => (
            <div
              key={i}
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: "4px",
              }}
            >
              <div style={{ display: "flex", alignItems: "center" }}>
                <Icon iconName="Attach" style={{ fontSize: 14, marginRight: "6px" }} />
                <span style={{ fontSize: "12px" }}>{file.name}</span>
              </div>
              <button
                onClick={() => removeFile(i)}
                style={{
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  color: "#a4262c",
                  fontSize: "12px",
                }}
              >
                ×
              </button>
            </div>
          ))}
        </div>
      )}

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
          marginTop: "10px",
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
