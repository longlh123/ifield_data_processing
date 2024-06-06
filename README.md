ABC
LONG
<h3>iField data processing</h3>

<blockquote>
<p><b>Step 1</b>: To see local branch</p>

<pre>
    <code>
        git branch
    </code>
</pre>

<p><b>Step 2</b>: To create and move to a new local branch</p>

<pre>
    <code>
        git branch -b my-branch-name
    </code>
</pre>

<u>Note</u>: replace <b>my-branch-name</b> with whatever name you want

<p><b>Step 3</b>: Setup config.json</p>

<ol>
    <li><b>project_name</b>: Enter a project name</li>
    <li><b>run_mdd_source</b>: Select TRUE to create an original mdd/ddf file or FALSE to skip this step.</li>
</ol>

<p><b>Step 4</b>: Setup for the F2F/CLT/CATI (optional)</p>

<ol>
    <li><b>default_language: Default the language that is used as the default</b></li>
    <li><b>delete_all</b>: Delete all data before inserting new data (default is FALSE)</li>
    <li><b>dummy_data_required</b>: Allow inserting dummy data (default is FALSE)
    <li><b>remove_all_ids</b>: Remove all IDs with types of cancel and extra (default is TRUE)
</ol>

<p>Step 3: Setup for the F2F/CLT/CATI</p>

<ol>
    <li>
        <b>main</b>:
        <ul>
            <li>
                <b>xmls</b>: Fill in the information for the project's XML files (from the old version to the new version) according to the syntax:
                <pre>
                    <code>
                        "protodid": "file_name.xml"
                    </code>
                </pre>
            </li>
            <li><b>protoid_final</b>: Protoid is used fo create the file original mdd/ddf file</li>
        <ul>
    </li>
</ol>
</blockquote>

<pre>
    <code>
        {
            "project_name" : "--Enter a project name--",
            "run_mdd_source" : true, 
            "main" : {
                "xmls" : {
                    "--Enter a protodid 1--" : "--Enter a xml file--",
                    "--Enter a protodid 2--" : "--Enter a xml file--"
                },
                "protoid_final" : "--Enter a protoid final 2--"
            },
            "stages" : {
                "stage-1" : {
                    "xmls" : {
                        "--Enter a protodid 1--" : "--Enter a xml file--",
                        "--Enter a protodid 2--" : "--Enter a xml file--"
                    },
                    "protoid_final" : "--Enter a protoid final 2--"
                },
                "stage-2" : {
                    "xmls" : {
                        "--Enter a protodid 1--" : "--Enter a xml file--",
                        "--Enter a protodid 2--" : "--Enter a xml file--"
                    },
                    "protoid_final" : "--Enter a protoid final 2--"
                },
                "stage-3" : {
                    "xmls" : {
                        "--Enter a protodid 1--" : "--Enter a xml file--",
                        "--Enter a protodid 2--" : "--Enter a xml file--"
                    },
                    "protoid_final" : "--Enter a protoid final 2--"
                }
            }
        }
    </code>
</pre>