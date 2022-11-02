#!/bin/sh
# This is a helper script to take help text from text file and turn it into the
# required C# code.

begintext=$(cat <<'EOF'
using System.Text;
namespace excelchop;
public static class HelpText
{
    public static string Text()
    {
        StringBuilder sb = new StringBuilder();
EOF
)

endtext=$(cat <<'EOF'
        return sb.ToString();
    }
}

EOF
)

lines=$(while read -r line; do printf "        %s\"%s\"%s\n" 'sb.AppendLine(' "$line" ');' ; done < helptext.md)

printf "%s\n%s\n%s\n" "$begintext" "$lines" "$endtext"
