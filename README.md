<h3>iField data processing</h3>

<blockquote>
<p>Step 1: Setup config.json</p>

<ol>
    <li><b>project_name</b>: Enter a project name</li>
    <li><b>run_mdd_source</b>: Select TRUE to create an original mdd/ddf file or FALSE to skip this step.</li>
</ol>

<p>Step 2: Setup for the F2F/CLT/CATI</p>

<ol>
    <li>
        <b>main</b>:
        <ul>
            <li><b>xmls</b>: Điền thông tin các file xmls của dự án (từ version cũ -> version mới) theo cấu trúc <b>protodid: file_name.xml</b></li>
            <li><b>protoid_final</b>: Protoid được dùng để tạo file original mdd/ddf</li>
            <li><b>csv</b>: Điền thông tin file finalize csv</li>
        <ul>
    </li>
</ol>
</blockquote>

<pre>
    <code>
        {
            "project_name" : "###--Enter a project name--###",
            "run_mdd_source" : true, 
            "main" : {
                "xmls" : {
                    "###--Enter a protodid--###" : "###--Enter a xml file--###",
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