/** Windows command lines (no VS Code dependency, so it can be tested with plain Node). */

/** Splits a command line like Windows does (CommandLineToArgvW rules for quotes and backslashes). */
export function splitCommandLine(commandLine: string): string[] {
    const args: string[] = [];
    let current = "";
    let inQuotes = false;
    let hasToken = false;

    for (let i = 0; i < commandLine.length; i++) {
        const c = commandLine[i];

        if (c === "\\") {
            let slashes = 0;

            while (commandLine[i] === "\\") {
                slashes++;
                i++;
            }

            if (commandLine[i] === "\"") {
                current += "\\".repeat(Math.floor(slashes / 2));

                if (slashes % 2 === 1) {
                    current += "\"";
                } else {
                    i--;   // the quote is processed normally
                }
            } else {
                current += "\\".repeat(slashes);
                i--;
            }

            hasToken = true;
            continue;
        }

        if (c === "\"") {
            if (inQuotes && commandLine[i + 1] === "\"") {
                current += "\"";   // "" inside quotes
                i++;
            } else {
                inQuotes = !inQuotes;
            }

            hasToken = true;
            continue;
        }

        if ((c === " " || c === "\t") && !inQuotes) {
            if (hasToken) {
                args.push(current);
                current = "";
                hasToken = false;
            }

            continue;
        }

        current += c;
        hasToken = true;
    }

    if (hasToken) {
        args.push(current);
    }

    return args;
}


/** Quotes an argument for cmd.exe / the Windows command line when needed. */
export function quoteArgument(argument: string): string {
    if (argument.length > 0 && !/[\s"&|<>^()]/.test(argument)) {
        return argument;
    }

    return "\"" + argument.replace(/(\\*)"/g, "$1$1\\\"").replace(/(\\+)$/, "$1$1") + "\"";
}
