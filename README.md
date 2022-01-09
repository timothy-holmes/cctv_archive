# cctv_archive
Source code excerpts from VBA-based application to assist with CCTV condition monitoring footage archival. 

The job that no one wants to do is given to the new engineering graduate each year. Thousands of new CCTV jobs, gigabytes and gigabytes, need to be filed away according to the pipe segments they monitor. The archival method was non-negotiable, but there's SOO. MANY. FILES. To automate, I broke the prcoess into two parts, the search and the move, giving the operator (sadly, me) the opportunity to approve file operation queue before it's executed. And logging, berbose logging the file operations in case of an error discovered later.

- Class modules in VBA to deal with memory leakage
- Expert system to assist with common errors in labelling by CCTV contractor, including Levenshtein distance to make suggestions on near matches

Things I learned:
- Debugging from within classes is hard (used recursive error handling to pinpoint exceptions)
- VBA is similar to C#
- Getting approval from the IT department to run my script on their system would have required a bit of bullshitting (instead I did extensive testing)

```vb
'check perceived type by file extension according to registry
strPerceivedType = objWScriptShell.RegRead("HKCR\." & strExt & "\PerceivedType")
```
