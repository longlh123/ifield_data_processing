<h3>iField data processing</h3>

<blockquote>
<p>Step 1: Setup config.json</p>

<ol>
    <li>project_name: Enter a project name</li>
    <li>run_mdd_source: Select TRUE to create an original mdd/ddf file or FALSE to skip this step.</li>
</ol>
</blockquote>

<pre>
    <code>
        {
            "project_name" : "###--Enter a project name--###",
            "run_mdd_source" : true, 
            "main" : {
                "xmls" : {
                    "###--Enter a protodid--###" : "###--Enter a xml file--###"
                },
                "protoid_final" : "8402",
                "csv" : "###--Enter a csv file--###"
            },
            "stages" : {
                "stage-1" : {
                    "xmls" : {
                        "###--Enter a protodid--###" : "###--Enter a xml file--###"
                    },
                    "protoid_final" : "8402",
                    "csv" : "###--Enter a csv file--###"
                },
                "stage-2" : {
                    "xmls" : {
                        "###--Enter a protodid--###" : "###--Enter a xml file--###"
                    },
                    "protoid_final" : "8402",
                    "csv" : "###--Enter a csv file--###"
                },
                "stage-3" : {
                    "xmls" : {
                        "###--Enter a protodid--###" : "###--Enter a xml file--###"
                    },
                    "protoid_final" : "8402",
                    "csv" : "###--Enter a csv file--###"
                }
            }
        }
    </code>
</pre>