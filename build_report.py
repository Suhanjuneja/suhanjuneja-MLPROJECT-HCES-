"""Generate the editable journal article from real saved experiment results.
Run after: python src/analysis.py --data data/CSV_data_HH_Cons_exp_22_23.zip
Output: paper/Suhan_Juneja_HCES_P54_Research_Paper.docx
For PDF use Word's File > Export > Create PDF (or installed LibreOffice).
"""
from pathlib import Path
import json
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parent
OUT=ROOT/'paper';OUT.mkdir(exist_ok=True)
R=ROOT/'results'; F=R/'figures'
report=json.loads((R/'run_summary.json').read_text(encoding='utf8'))
profile=pd.read_csv(R/'cluster_profiles.csv')
compare=pd.read_csv(R/'model_comparison.csv')
ward=pd.read_csv(R/'ward_comparison_same_sample.csv')
sector=pd.read_csv(R/'sector_cluster_percent.csv')
if report['best_k']!=5 or report['data']['eligible_households']!=261744:
    raise RuntimeError('Results changed. Please update interpretations in build_report.py before rebuilding.')

D=Document(); sec=D.sections[0];sec.top_margin=Inches(.78);sec.bottom_margin=Inches(.68);sec.left_margin=Inches(.88);sec.right_margin=Inches(.88)
sec.header_distance=Inches(.3);sec.footer_distance=Inches(.32)
styles=D.styles
normal=styles['Normal'];normal.font.name='Cambria';normal.font.size=Pt(10.1);normal.font.color.rgb=RGBColor(34,45,57)
normal.paragraph_format.space_after=Pt(5.5);normal.paragraph_format.line_spacing=1.09
for name,size,space in [('Title',16,10),('Heading 1',12.1,8),('Heading 2',10.7,5)]:
 st=styles[name];st.font.name='Aptos Display';st.font.size=Pt(size);st.font.bold=True;st.font.color.rgb=RGBColor(30,65,85)
 st.paragraph_format.space_before=Pt(space);st.paragraph_format.space_after=Pt(4)
if 'Figure caption' not in styles:
 st=styles.add_style('Figure caption',WD_STYLE_TYPE.PARAGRAPH)
else:st=styles['Figure caption']
st.font.name='Cambria';st.font.size=Pt(8.1);st.font.italic=True;st.font.color.rgb=RGBColor(73,88,99)
st.paragraph_format.space_after=Pt(10);st.paragraph_format.keep_with_next=False

header=sec.header.paragraphs[0];header.text='MACHINE LEARNING APPLICATION | ORIGINAL RESEARCH | HCES 2022–23';header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
header.runs[0].font.name='Aptos';header.runs[0].font.size=Pt(7);header.runs[0].font.bold=True;header.runs[0].font.color.rgb=RGBColor(104,115,127)
footer=sec.footer.paragraphs[0];footer.text='Suhan Juneja  •  Problem 54  •  Household-level consumption segmentation     |     Page '
footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
for r in footer.runs:r.font.size=Pt(7.3);r.font.color.rgb=RGBColor(90,100,112)
run=footer.add_run();begin=OxmlElement('w:fldChar');begin.set(qn('w:fldCharType'),'begin');run._r.append(begin)
run=footer.add_run('PAGE');instr=OxmlElement('w:instrText');instr.set(qn('xml:space'),'preserve');instr.text=' PAGE ';run._r.append(instr)
run=footer.add_run();end=OxmlElement('w:fldChar');end.set(qn('w:fldCharType'),'end');run._r.append(end)


def heading(t,level=1):D.add_heading(t,level=level)
def para(t, bold=False):
 p=D.add_paragraph();p.paragraph_format.widow_control=True
 r=p.add_run(t);r.bold=bold;return p

def bullets(lines):
 for line in lines:
  p=D.add_paragraph(style='List Bullet');p.add_run(line)

def table(headers,rows,widths=None,formats=None):
 t=D.add_table(rows=1, cols=len(headers));t.style='Light Shading Accent 1';t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for j,h in enumerate(headers):
  cell=t.rows[0].cells[j];cell.text=str(h)
  for run in cell.paragraphs[0].runs:run.bold=True;run.font.size=Pt(8);run.font.color.rgb=RGBColor(32,55,65)
 for row in rows:
  c=t.add_row().cells
  for j,v in enumerate(row):
   c[j].text=str(v);c[j].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
   for p in c[j].paragraphs:
    p.paragraph_format.space_after=Pt(1);p.paragraph_format.space_before=Pt(1)
    for run in p.runs:run.font.name='Cambria';run.font.size=Pt(8)
 for r in t.rows:
  for i,c in enumerate(r.cells):
   if widths and i<len(widths):c.width=Inches(widths[i])
 D.add_paragraph().paragraph_format.space_after=Pt(0)
 return t

fig_number=0
def fig(image,desc,width=6.1):
 global fig_number
 fig_number+=1
 p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(6);p.paragraph_format.space_after=Pt(1)
 p.paragraph_format.keep_with_next=True
 p.add_run().add_picture(str(F/image),width=Inches(width))
 cap=D.add_paragraph(style='Figure caption');cap.alignment=WD_ALIGN_PARAGRAPH.CENTER
 cap.add_run(f'Figure {fig_number}. {desc}')

# ===== Article front matter =====
t=D.add_paragraph(style='Title');t.alignment=WD_ALIGN_PARAGRAPH.CENTER
t.add_run('Household Consumption-Budget Segmentation in India Using HCES 2022–23 Unit-Level Microdata')
p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('Suhan Juneja');r.bold=True;r.font.size=Pt(11)
p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run('Department of Computer Science, Ramanujan College, University of Delhi, New Delhi, India\nRoll No. 24020570031  |  Machine Learning — B.Sc. (H) Computer Science  |  Allocated Problem 54')
r.font.size=Pt(8.4);r.font.color.rgb=RGBColor(80,92,104)
heading('Abstract',2)
para('This study identifies household consumption profiles from food and non-food budget shares. We analyze the supplied unit-level Household Consumption Expenditure Survey (HCES) 2022–23 data, joining household characteristics (Level 03) with questionnaire summary items (Level 14). A total of 8,075,528 summary-item records correspond to 261,746 sampled households; 261,744 have a positive reconstructed monthly consumption proxy and valid household metadata. Five input features represent the shares spent on cereals and pulses, milk and animal-source foods, served/packaged processed food, fuel and light, and selected services. The remaining expenditure forms an explicit “other” category. Values with seven-day and 365-day reference periods are converted to 30-day equivalents. K-means and MiniBatch k-means are compared across two to six segments using a fixed 45,000-household fitting sample, with silhouette scores computed on 2,100 selected observations. K-means at k = 5 achieves the highest primary comparison score (0.2216), indicating modest—not strong—separation. A separate Ward benchmark, repeated-subsample stability (mean adjusted Rand index 0.907), and a processed-food feature-ablation test (ARI 0.461) demonstrate that the segments are descriptive and sensitive to feature definitions. The work is fully scripted, and only aggregated results are published. It does not produce official MPCE, measure nutritional adequacy, or infer causality.')
para('Keywords: HCES 2022–23; household microdata; budget shares; K-means; hierarchical clustering; silhouette; stability; consumption segmentation',bold=True)

heading('1. Introduction')
para('Patterns of household spending describe how families allocate limited resources among competing needs. A household that devotes relatively more of its recorded consumption value to staple grains differs in expenditure composition from one with higher service, protein-source or convenience-food shares. In India, the Household Consumption Expenditure Survey (HCES) provides detailed information from separate food, consumables/services and durables questionnaires. The official statistical purpose includes consumption pattern analysis and monthly per-capita consumer expenditure estimation [1].')
para('This study addresses Problem 54 by investigating household-level consumption-pattern segmentation using cereals, protein foods, processed food, fuel and services. Unit-level HCES observations capture differences in the expenditure composition of individual sampled households, including households within the same State/UT. The objective is descriptive unsupervised learning, not prediction of poverty status or inference about nutritional health.')
para('Our research questions are: (RQ1) Can five interpretable household budget-share variables reveal distinguishable consumption profiles? (RQ2) How do candidate cluster counts and algorithms compare under an explicit internal quality criterion? (RQ3) Are the segments stable under resampling and omission of a short-reference-period food variable? (RQ4) Do their sample compositions differ between rural and urban respondents?')

heading('2. Background and Related Work')
para('The HCES 2022–23 design is a multistage stratified household sample. The survey used FDQ, CSQ and DGQ questionnaires in separate visits, with household characteristics collected in HCQ; consequently, values recorded over seven, thirty and 365 days require reconciliation before comparison [1, 2]. Official survey multipliers support population estimation, but the present exercise focuses on unweighted individual household composition for clustering and intentionally makes no State-level population estimate.')
para('K-means finds k centroids by minimizing within-cluster squared Euclidean distance and is a standard baseline for partitioning continuous variables [3]. The silhouette coefficient balances cohesion and separation of a point relative to neighboring groups, with higher positive values generally preferred [4]. Ward hierarchical clustering is a useful algorithmic contrast, while the adjusted Rand index (ARI) compares cluster assignments after accounting for chance; it is invariant to relabeling of arbitrary cluster IDs. Standardization is critical because the original five shares have unequal variability, and otherwise the largest-variance share can dominate Euclidean distances.')

heading('3. Dataset and Construction of Variables')
heading('3.1 Sources, household identifiers, and eligibility',2)
para('The original supplied archive contains fifteen CSV levels. This analysis deliberately uses two: Level 03 (household size and sample identifiers) and Level 14 (A1/B1/C1 summary item codes and values). The supplied layout workbook identifies the household linking fields. A six-field composite key—FSU, sector, State/UT, sample SU, second-stage stratum and sample household number—is unique across all 261,746 Level 03 household observations. Codes such as 01 and 1.0 are normalized consistently before joining. The Level 14 summary contains 3,149,267 food, 2,540,514 consumables/services and 2,385,747 durable summary records. There are no unmapped summary codes in this supplied archive under the Appendix A coding dictionary.')
para('The merge produces 261,746 households. The study excludes two households with no positive reconstructed total; 261,744 remain. Rows with absent item subtotals are interpreted as zero expenditure in that category, while missing or negative recorded values are never silently counted. The analysis package contains no individual household identifiers or row-level survey records. Figure 1 documents the selection count; the full code records these counts automatically.')
fig('01_coverage.png','Screening from matched summary records to eligible households. Two records were excluded because they had no positive reconstructed total.',6.0)

heading('3.2 Feature dictionary and correct recall-period conversion',2)
para('Table 1 records the exact summary item codes; these are section subtotals, not individual food-item entries. Using subtotals once avoids double-counting. Four figures are standardized per household by constructing a thirty-day proxy: a value observed over seven days is multiplied by 30/7, a 365-day value by 30/365, and a thirty-day value is unchanged. This practical conversion is described precisely because HCES field questionnaires mix reference periods [2].')
para('Table 1. Definition of budget-share inputs using the official A1/B1/C1 summary subtotals.',bold=True)
table(['Budget feature','Questionnaire / summary item codes','Interpretation'],[
 ['Staples and pulses','FDQ A1: 129, 139, 159','Cereals, cereal substitutes, pulses and products'],
 ['Protein-source foods','FDQ A1: 169, 199','Milk/products and egg, fish and meat (not pulse protein)'],
 ['Processed food','FDQ A1: 289, 299','Served and packaged processed food'],
 ['Fuel and light','CSQ B1: 349','Fuel and lighting expenditure'],
 ['Services','CSQ B1: 409, 419, 429, 519, 499, 439, 529','Education, medical, transport, services, entertainment and paid rent'],
 ['Other denominator','All remaining official A1/B1/C1 subtotals except code 539','Vegetables, fruits, oil, clothing, durables etc.'],
], [1.3,2.35,2.65])
para('For a selected category g and household h, the input share is s(h,g) = monthly_value(h,g) / sum(monthly_value(h,all included categories)). The “other” remainder is retained only in the denominator and visualization; the model fits five selected shares. Imputed house rent (code 539) is omitted, as distinct from actually paid rent (529). Household expenditure is measured at the household level, so dividing these shares by household size would change nothing and is unnecessary.')
para('This is a transparent project-specific consumption proxy, not the official published MPCE calculation. In particular, the study does not implement the full official cross-visit member-size adjustments, survey multiplier estimator or design-based standard errors. Hence no claims about India-wide spending levels or district estimates are made [1].')

heading('3.3 Data screening and exploratory distributions',2)
para('Budget shares are bounded between zero and one but are not normally distributed. For example, across the eligible sample the mean staples/pulses share is 6.9%, the mean protein-source-food share is 13.1%, and the mean services share is 26.0%. The unmodeled “other” category represents 38.4% on average. Median shares are often smaller than means because a subset of households exhibits much larger category shares. Figure 2 makes skew and occasional extremes visible instead of hiding them by deleting observations.')
fig('02_share_histograms.png','Histograms of the five modeled household budget shares across eligible survey records; all values are household-level proxies.',6.1)
fig('09_correlations.png','Pearson correlations among five shares. Since they share a denominator, negative associations can arise mechanically as well as behaviorally.',4.75)
para('Services share is negatively correlated with both staples and protein-source foods. These correlations are compositional—not causal—and should not be interpreted as food being displaced by services without further analysis of levels, household size, prices and survey context.')

heading('4. Methodology')
heading('4.1 Preprocessing and baseline models',2)
para('All experiments use random seed 42. To make the analysis reproducible on a typical student laptop, a fixed 45,000-household random subset is drawn without replacement to fit centroids. StandardScaler uses that fitting subset only, and the resulting transform and centroids assign every eligible household to a segment. The primary comparison is K-means versus MiniBatch k-means for k = 2, 3, 4, 5 and 6. MiniBatch k-means is a meaningful computational baseline because it approximates centroid updates from batches instead of using every observation in each iteration. Both methods are trained on the same subset with the same five features.')
para('For the selected objective, the K-means loss is J = Σ_i ||x_i − μ_{c(i)}||², where x_i is a standardized five-share vector and μ_{c(i)} is its assigned centroid. The selected segment count maximizes the mean silhouette on the SAME fixed 2,100-record evaluation subset, drawn from the 45,000-record fitting set. This is an internal unsupervised quality criterion, not a held-out prediction score or external accuracy.')
heading('4.2 Independent algorithmic benchmark and stability checks',2)
para('Ward agglomerative clustering is too expensive to fit to all sampled households in a beginner setting. We therefore run a separate, fair small-sample benchmark: both Ward and a fresh K-means fit use identical 2,100 standardized records for each k from two through six. These scores answer an algorithmic sensitivity question within that sample, not a head-to-head comparison of a full-data Ward model against 45,000-case K-means.')
para('For assignment stability, 12 paired random 8,000-record subsamples are independently clustered at the selected k. Each pair is compared using ARI computed only on household indices shared by both subsamples; a mean near one denotes relatively stable assignments on those overlaps. To evaluate variable dependence, the processed-food share is removed and k-means is refitted to the other four restandardized features. ARI between the original and reduced-feature fitting-sample assignments measures robustness to that particular feature choice, not inferential validity.')
heading('4.3 PCA, descriptive comparisons, and ethics',2)
para('Two-dimensional principal component analysis (PCA) is fitted to the standardized five-variable fitting sample solely for visualization, not for determining the clusters. Rural-versus-urban composition is summarized using raw sample proportions, not population-level statistics. A segment label is arbitrary and means no household has been classified as good, bad, wealthy, poor or nutritionally adequate. The code does not extract or export direct household identifiers, and the released graphics disclose anonymous aggregations or limited chart scatter only.')

heading('5. Experimental Results')
heading('5.1 Primary comparison and segment selection',2)
para('Table 2. Primary algorithm and cluster-count comparisons (same fitting and evaluation sets).',bold=True)
table(['Algorithm','k','Silhouette','Davies–Bouldin'],[
 [str(row.algorithm),int(row.k),f'{row.silhouette:.4f}',f'{row.davies_bouldin:.3f}']
 for _,row in compare.sort_values(['algorithm','k']).iterrows()
],[2.0,.5,1.2,1.4])
para('Table 2 compares ten primary model configurations. The highest silhouette is 0.2216 for standard K-means with five clusters; three clusters are close behind at 0.2161, a difference of just 0.0055. MiniBatch k-means performs somewhat worse in this experiment; its highest score is 0.2024 at two clusters. These scores are modest. There is structure in the feature space, but substantial overlap remains, and multiple segment counts could reasonably be used in a different application.')
fig('04_k_comparison.png','Primary comparison: K-means and MiniBatch k-means fitted on the same 45,000 cases and evaluated on the same 2,100-case subset.',6.1)
heading('5.2 Separate Ward benchmark',2)
para('When Ward and K-means are both refitted to precisely the same 2,100-record sample, Ward at k = 2 obtains a silhouette of 0.2249, slightly exceeding the K-means k = 5 silhouette of 0.2203 within this small benchmark. Ward scores fall as k increases. This finding warns against treating the five-cluster choice as a universal optimum; it was selected by the stated primary comparison protocol and is a useful descriptive, more granular solution.')
fig('12_ward_benchmark.png','Small-sample algorithmic contrast: Ward and K-means on identical 2,100 households, intentionally separate from the primary comparison.',6.1)

heading('5.3 Interpretation of the five fitted segments',2)
para('The following descriptions refer only to relative shares of reported consumption, not household income or dietary quality. Segment labels are assigned in descending mean staples-share order for readability after the unsupervised fit. In Table 3, the “other” share prevents readers from assuming that the five modeled categories exhaust all consumption.')
para('Table 3. Eligible households and mean category shares by segment (percent of reconstructed monthly total).',bold=True)
table(['Segment','Sample n','Share %','Staples','Protein','Processed','Fuel','Services','Other'],[
 [int(r.cluster),f'{int(r.sample_households):,}',f'{r.sample_pct:.1f}',f'{r.staples_share_mean_pct:.1f}',
  f'{r.protein_share_mean_pct:.1f}',f'{r.processed_share_mean_pct:.1f}',f'{r.fuel_share_mean_pct:.1f}',
  f'{r.services_share_mean_pct:.1f}',f'{r.other_share_mean_pct:.1f}'] for _,r in profile.iterrows()],
 [0.68,0.72,0.7,.7,.7,.8,.62,.72,.7])
para('Segment 1 (18.4% of sample, 48,024 households) is relatively staples-oriented: its average cereals-and-pulses share is 12.4%, versus 4.9% in Segment 4. Segment 2 (14.6%, 38,325 households) stands out for fuel and light: 14.0% versus 5.9% in Segment 4. Segment 3 (32.0%, 83,662 households) has the greatest milk-and-animal-source-food share at 19.4%. Segment 4 (32.3%, 84,490 households) is services-oriented, with 39.8% of its reconstructed budget proxy in the selected service categories. Segment 5 (2.8%, 7,243 households) is unusually processed-food-oriented, averaging 42.1% in served or packaged processed foods.')
fig('06_cluster_profiles.png','Mean budget shares in five fitted household segments; the contrast is about composition rather than total rupee spending.',6.15)
fig('07_budget_composition.png','Average share composition with the residual “other” category, showing why selected budget features do not sum to 100%.',6.15)
para('The small processed-food segment must be handled especially cautiously: 54.4% of its households have size one, versus 3.1% in Segment 3. A seven-day food reporting window may capture short-term circumstances rather than a stable lifestyle. This is a substantive limitation, not a reason to delete the group or portray it as a permanent type.')

heading('5.4 Rural/urban descriptive split and reduced-dimensional view',2)
para('Sample composition differs by sector, as visualized in Figure 8. In the rural sample, 24.0% of households are assigned to Segment 1 and 22.8% to Segment 4; in the urban sample, the corresponding shares are 10.2% and 46.1%. Segment 5 comprises 1.8% of rural and 4.1% of urban sample households. Because weights and stratification are not incorporated into these fractions, these are not estimates of rural/urban population percentages or statistically tested differences.')
fig('08_sector_segments.png','Distribution of model-assigned segments within rural and urban *sample* households; no population weighting is claimed.',6.05)
fig('03_rural_urban.png','Unweighted rural and urban mean category shares. Services account for a higher raw sample share among urban respondents.',6.05)
para('The first two principal components explain 34.8% and 24.3% of the standardized five-share feature variance (59.1% together). The two-dimensional plot therefore omits 40.9% of modeled variance. Overlap between colored point clouds is unsurprising and must not be described as evidence of clear separation or a supervised accuracy score.')
fig('05_pca.png','Two-dimensional PCA visualization of an anonymous 2,100-case sample; clustering used all five standardized shares.',5.8)

heading('6. Stability, Sensitivity, and Uncertainty')
para('Across 12 pairs of independently drawn 8,000-household subsamples, the adjusted Rand index (ARI) on overlapping observations averages 0.907, with a minimum of 0.817 and maximum of 0.962. This suggests reasonable initialization-and-subsampling robustness within a fixed five-feature representation. The diagnostic does not reproduce the complex survey sampling design, and overlapping subsamples are not independent new surveys.')
fig('10_stability.png','Adjusted Rand index for 12 paired 8,000-household subsamples, computed on observations included in both fits.',6.1)
para('The more demanding feature-ablation test is notably less reassuring: when processed-food share is dropped and clusters are recomputed at k = 5, agreement with the original labels is only ARI = 0.461. Thus, a plausible variable choice meaningfully changes assignments. Together with the modest silhouette and the strong presence of single-member households in Segment 5, this argues for interpreting the results as useful exploratory groupings, not a stable social or nutritional taxonomy.')
fig('11_segment_sizes.png','Unequal segment sizes; the processed-food-oriented segment contains fewer than 3% of analyzed households.',5.5)

heading('7. Discussion, Limitations, and Implications')
para('The household is the observational unit throughout the analysis. Using 261,744 eligible observations captures within-State heterogeneity, while official A1/B1/C1 subtotals permit interpretable expenditure-share features. The algorithm comparison demonstrates the importance of an explicit baseline: MiniBatch k-means did not automatically dominate standard K-means, and Ward favored fewer clusters on the matched smaller sample.')
para('Several limitations materially bound the interpretation. First, these are cross-sectional consumption-value proxies from mixed reference periods, not official MPCE. Second, the five features do not represent the full budget: other categories average 38.4%. Third, survey weights, clustering of first-stage sampling units, nonresponse adjustments and design-based uncertainty are not modeled; the descriptive sector percentages should not be generalized to households nationwide. Fourth, compositional features share a denominator, so correlations and distances are sensitive to relative share definitions. Fifth, food consumption values do not measure nutrients, calories, dietary quality or health outcomes. Sixth, major annual expenses such as education, paid rent and durables can be lumpy when expressed as monthly equivalents. Seventh, categorical household characteristics in HCQ have reporting limitations, and the project does not draw population social-group or district-level estimates [1, 2].')
para('Finally, because k was selected on the same fixed subset used for the internal silhouette comparison, the reported score is not an unbiased external generalization estimate. Different sample draws, feature sets, robust transformations or compositional log-ratio methods may change results. Ethical use requires explaining these uncertainties rather than turning cluster labels into judgments about households.')

heading('8. Conclusion and Future Work')
para('With the supplied unit-level HCES 2022–23 data, household expenditure composition can be segmented transparently using classical machine learning. Under the primary stated protocol, five-cluster K-means offers the highest tested silhouette (0.2216). The resulting groups emphasize staples/pulses, fuel, milk and animal-source foods, services, and processed foods, respectively, but overlap substantially. Resampling ARI is high (0.907 average), whereas removing processed-food share changes many assignments (ARI 0.461). The appropriate conclusion is the existence of interpretable exploratory patterns—not verified nutrition, poverty, income or permanent household categories.')
para('Further work should evaluate compositional data transformations, check alternative time-window normalizations, perform survey-weight-aware descriptive estimation where methodologically justified, replicate clustering on independent HCES rounds, and assess how stable interpretations are when unusual single-person and high-processed-share records are treated separately. The scripts, parameter choices and aggregated outputs in the repository make these extensions testable.')

heading('Data Availability, Reproducibility, and Declarations')
para('Data and code availability. This study uses the HCES 2022–23 unit-level survey files and accompanying MoSPI/NSSO technical documentation. The supplementary project includes the analysis scripts, input-file specification (data/README.md), aggregated output tables, diagnostic figures and the recorded run parameters in results/run_summary.json. Reproducibility: install the dependencies in requirements.txt, provide the documented survey input file, run `python src/analysis.py`, and then run `python build_report.py` to regenerate the editable manuscript. Feature definitions and random seeds are specified in src/analysis.py.')
para('Conflicts of interest and funding. No conflict of interest or external funding is declared. AI assistance. AI tools assisted with research writing, code development and figure preparation; the reported experiments were run against the HCES source files.')

heading('References')
refs=[
 '[1] Ministry of Statistics and Programme Implementation (MoSPI), National Sample Survey Office, “Household Consumption Expenditure Survey 2022–23: Survey Methodology and Estimation Procedure,” supplied methods note, pp. 1–13.',
 '[2] Ministry of Statistics and Programme Implementation (MoSPI), “HCES 2022–23 Questionnaires, Appendix A,” survey questionnaire, Sections A1/B1/C1. https://www.mospi.gov.in/sites/default/files/publication_reports/HCES-22-23/AppendixA.pdf',
 '[3] J. MacQueen, “Some Methods for Classification and Analysis of Multivariate Observations,” Proc. Fifth Berkeley Symp. Mathematical Statistics and Probability, 1967, pp. 281–297.',
 '[4] P. J. Rousseeuw, “Silhouettes: A Graphical Aid to the Interpretation and Validation of Cluster Analysis,” Journal of Computational and Applied Mathematics, 20, 53–65, 1987. https://doi.org/10.1016/0377-0427(87)90125-7',
 '[5] L. Hubert and P. Arabie, “Comparing Partitions,” Journal of Classification 2, 193–218, 1985. https://doi.org/10.1007/BF01908075',
 '[6] F. Pedregosa et al., “Scikit-learn: Machine Learning in Python,” Journal of Machine Learning Research, 12, 2825–2830, 2011. https://jmlr.org/papers/v12/pedregosa11a.html',
]
for x in refs:
 p=D.add_paragraph(x);p.paragraph_format.space_after=Pt(2)
 for r in p.runs:r.font.size=Pt(8.4)

out=OUT/'Suhan_Juneja_HCES_P54_Research_Paper.docx'
D.save(out)
print('Wrote',out)
