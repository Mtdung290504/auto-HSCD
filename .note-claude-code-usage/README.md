Sử dụng claude code qua API của AgentRouter

Nếu chưa cài claude-code, chạy:

```bash
npm install -g @anthropic-ai/claude-code
```

tại thư mục `C:\Users\<username>\.claude\settings.json`, điền:

```json
{
  "env": {
    "ANTHROPIC_BASE_URL": "https://agentrouter.org",
    "ANTHROPIC_AUTH_TOKEN": "YOUR_API_KEY",
    "ANTHROPIC_MODEL": "claude-opus-4-8"
  }
}
```

Vào thư mục workplace cần làm việc, mở terminal chạy:

```bash
claude.cmd # Trong trường hợp new chat
claude.cmd -r # Để hiện danh sách session cũ để chọn
```
