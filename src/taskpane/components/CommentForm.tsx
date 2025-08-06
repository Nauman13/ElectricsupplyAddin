import * as React from "react";
import { useState, useEffect } from "react";
import { MentionsInput, Mention } from "react-mentions";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./msalConfig";

// SharePoint Config
const siteId = "8314c8ba-c25a-4a02-bf25-d6238949ac8f";
const listId = "5f59364d-9808-4d26-8e04-2527b4fc403e";

const msalInstance = new PublicClientApplication(msalConfig);

const CommentForm: React.FC = () => {
  const [comment, setComment] = useState<string>("");
  const [commentHistory, setCommentHistory] = useState<any[]>([]);
  const [people, setPeople] = useState<any[]>([]);
  const [mentionedEmails, setMentionedEmails] = useState<string[]>([]);
  const [conversationId, setConversationId] = useState<string>("");

  const waitForMailboxItem = (): Promise<void> => {
    return new Promise((resolve) => {
      const check = () => {
        if (Office.context?.mailbox?.item) {
          resolve();
        } else {
          setTimeout(check, 100); // keep checking every 100ms
        }
      };
      check();
    });
  };

  useEffect(() => {
    Office.onReady(async (info) => {
      if (info.host === Office.HostType.Outlook) {
        await waitForMailboxItem();
        const item = Office.context.mailbox.item;

        item.body.getAsync(Office.CoercionType.Html, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const bodyContent = result.value;
            const match = bodyContent.match(/CONVERSATION_ID:([a-zA-Z0-9\-]+)/);

            let convId = match?.[1] || item.conversationId;

            if (convId) {
              setConversationId(convId);
              await fetchCommentsFromSharePoint(convId);
            }
          } else {
            console.error("Failed to get body:", result.error);
          }
        });
      }
    });
  }, []);

  // useEffect(() => {
  //   if (conversationId) fetchCommentsFromSharePoint();
  // }, [conversationId]);

  useEffect(() => {
    fetchUsers();
  }, []);

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

  const forwardOriginalEmailToMentionedUsers = async () => {
    if (mentionedEmails.length === 0) return;

    const token = await getAccessToken();

    const encodedItemId = Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    try {
      // Get message from Graph
      const messageRes = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${encodedItemId}`,
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );

      if (!messageRes.ok) {
        console.error("Failed to get original message:", await messageRes.text());
        return;
      }

      const message = await messageRes.json();
      const graphMessageId = message.id;

      // Forward message to each mentioned user with a notification comment
      for (const email of mentionedEmails) {
        await fetch(`https://graph.microsoft.com/v1.0/me/messages/${graphMessageId}/forward`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            comment: `Hi, you were mentioned in a comment on this email. Please review the original message below.`,
            toRecipients: [
              {
                emailAddress: {
                  address: email,
                },
              },
            ],
          }),
        });
        console.log(`Forwarded to ${email}`);
      }
    } catch (err) {
      console.error("Error forwarding email:", err);
    }
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

    const token = await getAccessToken();
    const emailIdValue = encodeURIComponent(id);

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$filter=fields/EmailID eq '${emailIdValue}'&$orderby=createdDateTime asc`;

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      console.error("Fetch comments failed:", await response.text());
      return;
    }

    const data = await response.json();
    const comments = data.value
      .map((item: any) => item.fields)
      .sort((a, b) => new Date(a.CreatedDate).getTime() - new Date(b.CreatedDate).getTime());

    setCommentHistory(comments);
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

    await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ fields: fieldsData }),
    });

    setMentionedEmails(emails);
  };

  // const sendEmailToMentionedUsers = async () => {
  //   if (mentionedEmails.length === 0) return;

  //   const token = await getAccessToken();

  //   Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
  //     if (result.status === Office.AsyncResultStatus.Succeeded) {
  //       const originalBody = result.value;

  //       const toRecipients = mentionedEmails.map((email) => ({
  //         emailAddress: { address: email },
  //       }));

  //       const emailPayload = {
  //         message: {
  //           subject: "You’ve been mentioned in an Outlook conversation",
  //           body: {
  //             contentType: "HTML",
  //             content: `
  //               <p>Hello,</p>
  //               <p>You were mentioned in a conversation. Here's the original email:</p>
  //               <hr />
  //               ${originalBody}
  //               <p>Open your Outlook Add-in to reply or view further comments.</p>
  //             `,
  //           },
  //           toRecipients,
  //         },
  //         saveToSentItems: true,
  //       };

  //       await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
  //         method: "POST",
  //         headers: {
  //           Authorization: `Bearer ${token}`,
  //           "Content-Type": "application/json",
  //         },
  //         body: JSON.stringify(emailPayload),
  //       });
  //     }
  //   });
  // };

  const handleSaveAndShare = async () => {
    if (!comment.trim()) {
      alert("Please add a comment before saving.");
      return;
    }

    await saveCommentToSharePoint();
    setTimeout(() => fetchCommentsFromSharePoint(), 1000);

    if (mentionedEmails.length > 0) {
      await forwardOriginalEmailToMentionedUsers();
    }

    setComment("");
    setMentionedEmails([]);
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
        {commentHistory.length > 0 ? (
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

              <div>
                <div style={{ fontWeight: 600 }}>{c.CreatedBy}</div>
                <div style={{ margin: "5px 0" }}>{c.Comment}</div>
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
      </div>

      <button
        onClick={handleSaveAndShare}
        style={{
          backgroundColor: "#0078D4",
          color: "#fff",
          border: "none",
          padding: "10px 16px",
          borderRadius: "5px",
          float: "right",
        }}
      >
        Send
      </button>
    </div>
  );
};

export default CommentForm;
