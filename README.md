ConvertRTCWorkItemsToPPT
===

When working with Rational Team Concert for planning it's easy to produce lots of work items and then find it difficult to visualise how they are spread across the team.  This little app is intended to consume a csv file describing all the work items (easy produced from RTC via export to csv from a work item query) and to produce a diagram in power point of the various work items.

Over time it might do more automatic layout, but as a very first start, getting a ppt with all the right boxes in it for the user to produce their own layout is a start.

The font used is deliberately small as this sort of nonsense is only required when you have a larger team or lots of work items (and they all want to fit onto a single slide).

Assumptions
---

You're going to open a csv file with four columns
1  - ID
2 - summary
3 - owner
4 - story point cost

There should be no header row in the csv file.

This app assumes that excel and powerpoint are installed on your machine (developed using office 2016). YMMV on other versions.
