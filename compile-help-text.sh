# This is a helper script to take help text from text file and turn it into the
# required C# code.

cat <<'EOF'
using System.Text;

namespace excelconvert
{
    public static class HelpText
    {
        public static string Text()
        {
            StringBuilder sb = new StringBuilder();
EOF


endtext=$(cat <<'EOF'
            return sb.ToString();
        }
    }
}

EOF
)

while read line; do printf "            %s\"%s\"%s\n" 'sb.AppendLine(' "$line" ');' ; done < helptext.md

#printf "%s%s%s" "$begintext" "lines" "endtext" >> excelconvert\HelpText.cs
printf "%s\n%s\n%s\n" "$begintext" "$lines" "$endtext" 
