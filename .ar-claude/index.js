import { spawn } from "node:child_process";

const CONFIG = {
    baseUrl: "https://agentrouter.org",
    model: "claude-opus-4-8",
};

function readSecret(prompt) {
    return new Promise((resolve, reject) => {
        const stdin = process.stdin;
        let input = "";

        process.stdout.write(prompt);

        stdin.setRawMode(true);
        stdin.resume();
        stdin.setEncoding("utf8");

        function onData(char) {
            // Enter
            if (char === "\r" || char === "\n") {
                stdin.setRawMode(false);
                stdin.pause();
                stdin.removeListener("data", onData);

                process.stdout.write("\n");
                resolve(input);
                return;
            }

            // Ctrl+C
            if (char === "\u0003") {
                stdin.setRawMode(false);
                stdin.pause();
                stdin.removeListener("data", onData);

                process.stdout.write("\n");
                reject(new Error("Cancelled by user"));
                return;
            }

            // Backspace
            if (char === "\u007f" || char === "\b") {
                if (input.length > 0) {
                    input = input.slice(0, -1);
                }
                return;
            }

            input += char;
        }

        stdin.on("data", onData);
    });
}

async function main() {
    let apiKey;

    try {
        apiKey = await readSecret("AgentRouter API Key: ");
        if (!apiKey.trim())
            throw new Error("API key cannot be empty.");


        const env = {
            ...process.env,

            ANTHROPIC_AUTH_TOKEN: apiKey,
            ANTHROPIC_BASE_URL: CONFIG.baseUrl,
            ANTHROPIC_MODEL: CONFIG.model,
        };

        const claude = spawn(
            process.env.ComSpec || 'cmd.exe',
            ['/d', '/s', '/c', 'claude.cmd'],
            {
                stdio: 'inherit',
                env,
                windowsHide: false,
            },
        );

        claude.on("error", (error) => {
            console.error("\nFailed to start Claude Code:", error.message);
            process.exitCode = 1;
        });

        claude.on("exit", (code, signal) => {
            if (signal) {
                process.exitCode = 1;
            } else {
                process.exitCode = code ?? 0;
            }
        });
    } finally {
        // Xóa reference tới API key khỏi launcher sau khi spawn process con.
        apiKey = undefined;
    }
}

main().catch((error) => {
    console.error(`\n${error.message}`);
    process.exitCode = 1;
});
