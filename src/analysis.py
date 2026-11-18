"""HCES 2022-23 household-level consumption segmentation, Problem 54.

Run: python src/analysis.py --data data/CSV_data_HH_Cons_exp_22_23.zip
Input archive stays local. Individual household records are never written to the output.
Source: HCES questionnaire Appendix A and supplied Layout_HCES_2022_23_1.xlsx.
"""
from __future__ import annotations
import argparse, hashlib, json, os, sys, time, zipfile
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans, MiniBatchKMeans, AgglomerativeClustering
from sklearn.decomposition import PCA
from sklearn.metrics import adjusted_rand_score, silhouette_score, davies_bouldin_score
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

SEED=42
KEY=['fsu','sector','state','b1q1pt7','b1q1pt11','b1q1pt12']
SHARE=['staples','protein','processed','fuel','services']
NAMES=['Staples & pulses','Milk, eggs, fish, meat','Processed foods','Fuel & light','Services*']
# Codes in Appendix A, Sections A1/B1/C1. All codes below are section SUBTOTALS,
# not individual item codes, to prevent counting items together with their subtotal.
FOOD30={129,139,159,179}
FOOD7={169,219,239,249,199,189,269,279,289,299}
CSQ30={349,459,479,429,519,499,439,529,539,309,319,329}
CSQ365={409,419,899}
DGQ365={379,399,389,629,609,99,619,599,579,559,569,639,649}
ALL_CODES=FOOD30|FOOD7|CSQ30|CSQ365|DGQ365
EXCLUDE={539}  # Official survey does not include imputed house rent in MPCE.
GROUP_CODES={
    'staples':{129,139,159},
    'protein':{169,199},
    'processed':{289,299},
    'fuel':{349},
    'services':{409,419,429,519,499,439,529},
}
CODE_BUCKET={c:g for g,cs in GROUP_CODES.items() for c in cs}
MONTH_FACTOR={c:(30/7 if c in FOOD7 else 30/365 if c in CSQ365|DGQ365 else 1) for c in ALL_CODES}
# e.g. Section 8 fuel = 30-day; 7-day food * (30/7); 365-day education * (30/365).


def member(z, level):
    m=[s for s in z.namelist() if s.lower().endswith('.csv') and ('level - '+str(level).zfill(2)) in s.lower()]
    if len(m)!=1: raise ValueError(f'ZIP must contain one LEVEL - {level} file; found {m}')
    return m[0]


def clean_key(d):
    for c in KEY:
        d[c]=pd.to_numeric(d[c],errors='coerce').fillna(-1).astype('int32')
    return d


def extract(zip_path, chunksize=450_000):
    """Read HCQ level 03 once and A1/B1/C1 level 14 in chunks."""
    start=time.time()
    with zipfile.ZipFile(zip_path) as z:
        with z.open(member(z,'03')) as fh:
            hh=pd.read_csv(fh,usecols=KEY+['b2q2pt1','b4q4pt10','mult'],dtype=str)
        clean_key(hh)
        dup=int(hh.duplicated(KEY,keep=False).sum())
        if dup: raise ValueError(f'Household key not unique in HCQ Level 03: {dup}')
        hh['size']=pd.to_numeric(hh['b2q2pt1'],errors='coerce')
        hh['weight']=pd.to_numeric(hh['mult'],errors='coerce')
        hh=hh.drop(columns=['b2q2pt1','mult','b4q4pt10'])
        summary=[]; rows=0; invalid=0; unknown={}; q_counts={'F':0,'C':0,'D':0}
        with z.open(member(z,'14')) as fh:
            iterator=pd.read_csv(fh,usecols=KEY+['questionaire_no','ba1b1c1_2','ba1b1c1_3'],chunksize=chunksize,low_memory=False)
            for ix,part in enumerate(iterator,1):
                rows+=len(part); clean_key(part)
                part['code']=pd.to_numeric(part['ba1b1c1_2'],errors='coerce').fillna(-1).astype('int32')
                part['value']=pd.to_numeric(part['ba1b1c1_3'],errors='coerce')
                for q,n in part['questionaire_no'].value_counts().items(): q_counts[q]=q_counts.get(q,0)+int(n)
                # Detect non-reference codes instead of silently counting them.
                for k,v in part.loc[~part['code'].isin(ALL_CODES),'code'].value_counts().items():
                    unknown[str(k)]=unknown.get(str(k),0)+int(v)
                part=part[part['code'].isin(ALL_CODES-EXCLUDE)].copy()
                invalid+=int(((part['value']<0)|(~np.isfinite(part['value']))).sum())
                part=part[part['value'].ge(0)&part['value'].notna()]
                part['monthly']=part['value']*part['code'].map(MONTH_FACTOR).astype('float32')
                part['bucket']=part['code'].map(CODE_BUCKET).fillna('other')
                # Avoid persisting household-level rows; aggregate within chunk.
                gp=part.groupby(KEY+['bucket'],sort=False,observed=True)['monthly'].sum().reset_index()
                summary.append(gp)
                if ix%6==0: print(f'  Read {rows:,} A1/B1/C1 summary rows ...',flush=True)
    a=pd.concat(summary,ignore_index=True)
    a=a.groupby(KEY+['bucket'],sort=False,observed=True)['monthly'].sum().unstack('bucket',fill_value=0).reset_index()
    for g in SHARE+['other']:
        if g not in a: a[g]=0.
    a['total']=a[SHARE+['other']].sum(axis=1)
    # An observed positive total is mandatory; missing categories treated as zero only
    # when absent from the official summary rows (not unvalued item records).
    alln=len(a)
    merged=hh.merge(a,on=KEY,how='inner',validate='one_to_one')
    match=len(merged)
    merged=merged[(merged['size'].between(1,30))&(merged['weight']>0)&(merged['total']>0)].copy()
    for g in SHARE: merged[g+'_share']=merged[g]/merged['total']
    merged['other_share']=merged['other']/merged['total']
    merged['sector_name']=np.where(merged['sector']==1,'Rural','Urban')
    meta={'hcq_households':len(hh),'duplicated_hcq_keys':dup,'level14_rows':rows,
          'level14_questionnaire_rows':q_counts,'level14_unknown_code_rows':unknown,
          'invalid_or_missing_value_rows':invalid,'summary_households':alln,
          'matched_households':match,'eligible_households':len(merged),
          'excluded_after_join':match-len(merged),'elapsed_read_seconds':round(time.time()-start,1),
          'data_zip_size_bytes':zip_path.stat().st_size}
    print(f'Matched={match:,}; retained={len(merged):,}; read elapsed {meta["elapsed_read_seconds"]} s',flush=True)
    return merged,meta


def model_fit(raw, out, meta, *, max_n=45000, eval_n=2100):
    """Fit unsupervised clusterings on a seeded random household sample.

    A capped fitting sample makes Ward benchmark and silhouette computations
    practical on a beginner laptop; sample is chosen BEFORE computing clusters.
    KMeans & MiniBatch use same fit sample; Ward applied on same evaluation subsample
    to provide a transparent comparison rather than silently comparing sizes.
    """
    rng=np.random.default_rng(SEED)
    n=len(raw)
    selection=np.sort(rng.choice(n,min(max_n,n),replace=False))
    X=raw.iloc[selection].copy()
    feature_cols=[g+'_share' for g in SHARE]
    scaler=StandardScaler()
    xx=scaler.fit_transform(X[feature_cols].to_numpy(dtype=float))
    eval_ids=np.sort(rng.choice(len(X),min(eval_n,len(X)),replace=False))
    xe=xx[eval_ids]
    compare=[]; chosen={}
    for k in range(2,7):
        # Exactly same held-out evaluation subset for each choice. Internal metric,
        # NOT a supervised test set or external clustering validation.
        m=KMeans(n_clusters=k,n_init=12,random_state=SEED)
        labels=m.fit_predict(xx)
        mb=MiniBatchKMeans(n_clusters=k,n_init=8,batch_size=2048,random_state=SEED)
        lb=mb.fit_predict(xx)
        evlabels=labels[eval_ids]
        for name, lab, model in [('K-means',labels,m),('MiniBatch k-means',lb,mb)]:
            pred=lab[eval_ids]
            sc=float(silhouette_score(xe,pred,random_state=SEED)) if len(np.unique(pred))==k else np.nan
            db=float(davies_bouldin_score(xe,pred)) if len(np.unique(pred))==k else np.nan
            compare.append({'algorithm':name,'k':k,'silhouette':round(sc,5),'davies_bouldin':round(db,5),
                            'inertia_fit_sample':round(float(model.inertia_),3)})
            chosen[(name,k)]=(lab,model)
        print('  Compared k =',k,flush=True)
    comparison=pd.DataFrame(compare).sort_values(['silhouette'],ascending=False)
    comparison.to_csv(out/'model_comparison.csv',index=False)
    # Ward is fitted on the SAME 2,100-household subset as a K-means benchmark;
    # this is a separate comparison, not a claim of full-sample Ward fitting.
    ward_rows=[]
    for kk in range(2,7):
        kl=KMeans(n_clusters=kk,random_state=SEED,n_init=12).fit_predict(xe)
        wl=AgglomerativeClustering(n_clusters=kk,linkage='ward').fit_predict(xe)
        ward_rows.extend([
          {'algorithm':'K-means (2,100 subset)','k':kk,'silhouette':round(float(silhouette_score(xe,kl)),5)},
          {'algorithm':'Ward hierarchical (2,100 subset)','k':kk,'silhouette':round(float(silhouette_score(xe,wl)),5)}])
    pd.DataFrame(ward_rows).to_csv(out/'ward_comparison_same_sample.csv',index=False)
    best=comparison.iloc[0]
    algorithm=str(best['algorithm']); k=int(best['k'])
    final_labels,final_model=chosen[(algorithm,k)]
    # For every household, centroids assign the closest cluster; reflects training-sample structure.
    raw_all=np.asarray(raw[feature_cols].to_numpy(),dtype=np.float64)
    all_scaled=scaler.transform(raw_all)
    all_pred=final_model.predict(all_scaled)
    # Sort arbitrary labels in decreasing mean staples-share for interpretability.
    selected=X.copy(); selected['cluster_original']=final_labels
    selected['cluster']=selected['cluster_original']
    ranks=selected.groupby('cluster_original')['staples_share'].mean().sort_values(ascending=False).index.tolist()
    label_map={old:i+1 for i,old in enumerate(ranks)}
    all_new=np.array([label_map[v] for v in all_pred],dtype=np.int16)
    raw=raw.copy(); raw['cluster']=all_new
    selected['cluster']=selected['cluster_original'].map(label_map)
    # Internal sensitivity over multiple new subsamples, with exact overlap IDs
    # for the same-data-comparison; assigns each subsample's clusters on common cases.
    stability=[]
    sample_n=min(8000,len(X)); base_idx=np.arange(len(X))
    for trial in range(12):
        a_idx=np.sort(rng.choice(base_idx,sample_n,replace=False))
        b_idx=np.sort(rng.choice(base_idx,sample_n,replace=False))
        a_lab=KMeans(n_clusters=k,n_init=5,random_state=SEED+trial*2).fit_predict(xx[a_idx])
        b_lab=KMeans(n_clusters=k,n_init=5,random_state=SEED+trial*2+1).fit_predict(xx[b_idx])
        intersection,ia,ib=np.intersect1d(a_idx,b_idx,return_indices=True)
        stab=float(adjusted_rand_score(a_lab[ia],b_lab[ib]))
        stability.append({'comparison':trial+1,'shared_households':len(intersection),'ari':round(stab,5)})
    pd.DataFrame(stability).to_csv(out/'stability.csv',index=False)
    # State & rural-urban descriptive counts, never district-level estimates.
    profile=[]
    for cluster, part in raw.groupby('cluster'):
        row={'cluster':int(cluster),'sample_households':len(part),'sample_pct':100*len(part)/len(raw),
             'rural_sample_pct':100*(part['sector']==1).mean(),'median_household_size':float(part['size'].median())}
        for g in SHARE+['other']:
            row[g+'_share_mean_pct']=100*part[g+'_share'].mean()
            row[g+'_share_median_pct']=100*part[g+'_share'].median()
        profile.append(row)
    pro=pd.DataFrame(profile).sort_values('cluster')
    pro.to_csv(out/'cluster_profiles.csv',index=False)
    sector_tab=pd.crosstab(raw['sector_name'],raw['cluster'],normalize='index')*100
    sector_tab.to_csv(out/'sector_cluster_percent.csv')
    # Aggregate-only output: cluster counts by state/sector, suppress small cells.
    counts=raw.groupby(['state','sector_name','cluster']).size().rename('n').reset_index()
    counts['n']=counts['n'].where(counts['n']>=30,np.nan)
    counts.to_csv(out/'state_sector_cluster_counts_suppressed.csv',index=False)
    # Aggregate descriptive statistics; unweighted by design, population inference prohibited.
    desc=raw[[g+'_share' for g in SHARE+['other']]].describe(percentiles=[.10,.25,.5,.75,.90]).T
    desc.to_csv(out/'share_descriptives.csv')
    # PCA for visualization ONLY; clustering uses full five-share standardized representation.
    pca=PCA(n_components=2,random_state=SEED)
    projection=pca.fit_transform(xx)
    pca_explained=pca.explained_variance_ratio_.tolist()
    visual_data={'pca_x':projection[eval_ids,0], 'pca_y':projection[eval_ids,1],
                 'pca_cluster':selected.iloc[eval_ids]['cluster'].to_numpy()}
    # Feature-ablation diagnostic: recompute labels without processed-food share.
    # This estimates how dependent the segments are on the short 7-day category.
    reduced_cols=[g+'_share' for g in SHARE if g!='processed']
    sx=StandardScaler().fit_transform(X[reduced_cols].to_numpy(dtype=float))
    reduced_lab=KMeans(n_clusters=k,n_init=12,random_state=SEED).fit_predict(sx)
    ablation_ari=round(float(adjusted_rand_score(final_labels,reduced_lab)),5)
    # Check how atypical the highly processed-food segment is demographically.
    single_share={str(int(c)):round(float((v['size']==1).mean()),5) for c,v in raw.groupby('cluster')}
    # report details and limitations as structured run manifest
    report={'goal':'Problem 54: unsupervised household-level budget-share segmentation',
        'data_source':'MoSPI HCES 2022-23 supplied unit-level zip / Level 03 & Level 14',
        'not_official_mpce':True,'no_nutrition_or_causal_claims':True,
        'data':meta,'sampled_for_fitting':len(X),'seed':SEED,'input_features':feature_cols,
        'codes_in_denominator':sorted(ALL_CODES-EXCLUDE),'excluded_imputed_rent_code':539,
        'best_algorithm':algorithm,'best_k':k,'selection_metric':'highest sample silhouette on fixed 2100-household evaluation subset',
        'best_silhouette':round(float(best['silhouette']),5),'pca_variance':pca_explained,
        'ablation_excluding_processed_share_ari':ablation_ari,
        'single_person_share_by_cluster':single_share,
        'stability_ari_mean':round(float(pd.DataFrame(stability)['ari'].mean()),5),
        'stability_ari_min':round(float(pd.DataFrame(stability)['ari'].min()),5),
        'stability_ari_max':round(float(pd.DataFrame(stability)['ari'].max()),5),
        'cluster_counts':{str(int(k)):int(v) for k,v in raw['cluster'].value_counts().sort_index().items()},
        'cluster_share_columns':SHARE,'full_household_data_exported':False}
    (out/'run_summary.json').write_text(json.dumps(report,indent=2),encoding='utf8')
    return raw,selected,comparison,pro,sector_tab,desc,report,visual_data


def figures(raw,selected,comparison,profile,sector_tab,desc,report,visual_data,out):
    figs=out/'figures';figs.mkdir(exist_ok=True)
    plt.rcParams.update({'font.size':9,'axes.spines.top':False,'axes.spines.right':False})
    PAL=['#285c74','#d08752','#597f65','#8e70a1','#c1a357','#6e849c']
    def save(name):
        plt.tight_layout();plt.savefig(figs/name,dpi=185,bbox_inches='tight');plt.close()
    # 1 profile survey selection
    m=report['data'];fig,ax=plt.subplots(figsize=(7.3,2.5));tot=m['matched_households'];n=m['eligible_households'];
    ax.barh(['Matched HCQ + summaries'],n,color=PAL[0],label='Retained')
    ax.barh(['Matched HCQ + summaries'],tot-n,left=n,color='#d3d6dc',label='Excluded')
    ax.text(n*.47,0,f'{n:,} retained',color='white',ha='center',va='center',weight='bold')
    ax.set_xlim(0,tot*1.14);ax.set_xlabel('Household records');ax.legend(frameon=False,ncol=2,loc='lower center',bbox_to_anchor=(.5,-.45));save('01_coverage.png')
    # 2 category share distributions
    fig,axes=plt.subplots(1,5,figsize=(11,2.4),sharey=True)
    for g,ax in zip(SHARE,axes):
        ax.hist(raw[g+'_share']*100,bins=np.arange(0,102,4),color=PAL[0],alpha=.8)
        ax.set_title(g.title());ax.set_xlim(0,100);ax.set_xlabel('Budget share (%)')
    axes[0].set_ylabel('Households');save('02_share_histograms.png')
    # 3 compare mean share by sector
    table=raw.groupby('sector_name')[[g+'_share' for g in SHARE]].mean()*100
    fig,ax=plt.subplots(figsize=(8.5,3.2));x=np.arange(5);w=.35
    for i,sector in enumerate(['Rural','Urban']):ax.bar(x+(i-.5)*w,table.loc[sector].values,width=w,label=sector,color=PAL[i])
    ax.set_xticks(x,NAMES,rotation=10);ax.set_ylabel('Unweighted average share (%)');ax.legend(frameon=False);save('03_rural_urban.png')
    # 4 silhouettes for k
    fig,ax=plt.subplots(figsize=(7.3,3.25))
    for i,(name,sub) in enumerate(comparison.groupby('algorithm')):
        sub=sub.sort_values('k');ax.plot(sub.k,sub.silhouette,'o-',label=name,color=PAL[i],lw=2,ms=5)
    ax.set_xlabel('Cluster count (k)');ax.set_ylabel('Sample silhouette score');ax.set_xticks(range(2,7));ax.legend(frameon=False);save('04_k_comparison.png')
    # 5 PCA - sample only not a plot of raw microdata (each point anonymous)
    pic=visual_data
    fig,ax=plt.subplots(figsize=(6.7,4.2))
    for i in sorted(np.unique(pic['pca_cluster'])):
        mask=pic['pca_cluster']==i
        ax.scatter(pic['pca_x'][mask],pic['pca_y'][mask],s=8,alpha=.27,label=f'Segment {i}',color=PAL[(i-1)%len(PAL)],edgecolors='none')
    ax.set_xlabel(f'PC1 ({report["pca_variance"][0]*100:.1f}% variance)');ax.set_ylabel(f'PC2 ({report["pca_variance"][1]*100:.1f}% variance)');ax.legend(frameon=False,ncol=3,fontsize=8);save('05_pca.png')
    # 6 grouped profile bars
    fig,ax=plt.subplots(figsize=(9.2,3.4));x=np.arange(len(SHARE));w=.77/len(profile)
    for i,row in profile.reset_index().iterrows():
        ax.bar(x+(i-(len(profile)-1)/2)*w,[row[g+'_share_mean_pct'] for g in SHARE],width=w,label=f'Segment {int(row.cluster)}',color=PAL[i])
    ax.set_ylabel('Mean share (%)');ax.set_xticks(x,NAMES,rotation=8);ax.legend(frameon=False,ncol=min(4,len(profile)));save('06_cluster_profiles.png')
    # 7 100% stack budget (including other)
    fig,ax=plt.subplots(figsize=(8.5,3.2));bottom=np.zeros(len(profile))
    for i,g in enumerate(SHARE+['other']):
        v=profile[g+'_share_mean_pct'].to_numpy();ax.bar([f'Segment {c}' for c in profile.cluster],v,bottom=bottom,label=('Other' if g=='other' else g.title()),color=PAL[i]);bottom+=v
    ax.set_ylabel('Mean of household category shares (%)');ax.set_ylim(0,100);ax.legend(ncol=3,frameon=False,loc='lower center',bbox_to_anchor=(.5,-.44));save('07_budget_composition.png')
    # 8 sector mix stacked
    fig,ax=plt.subplots(figsize=(7,2.7));bottom=np.zeros(len(sector_tab))
    for i,c in enumerate(sector_tab.columns):
        v=sector_tab[c].to_numpy();ax.bar(sector_tab.index,v,bottom=bottom,color=PAL[(int(c)-1)%len(PAL)],label=f'Segment {int(c)}');bottom+=v
    ax.set_ylabel('Share of households in sector (%)');ax.set_ylim(0,100);ax.legend(frameon=False,ncol=3,loc='lower center',bbox_to_anchor=(.5,-.4));save('08_sector_segments.png')
    # 9 feature correlation heatmap
    mat=raw[[g+'_share' for g in SHARE]].corr();fig,ax=plt.subplots(figsize=(5.1,4.3));im=ax.imshow(mat,vmin=-1,vmax=1,cmap='RdBu_r');
    ax.set_xticks(range(5),[g.title() for g in SHARE],rotation=35,ha='right');ax.set_yticks(range(5),[g.title() for g in SHARE]);
    for i in range(5):
        for j in range(5):ax.text(j,i,f'{mat.iloc[i,j]:.2f}',ha='center',va='center',fontsize=8)
    fig.colorbar(im,ax=ax,shrink=.8,label='Pearson correlation');save('09_correlations.png')
    # 10 robustness
    stab=pd.read_csv(out/'stability.csv');fig,ax=plt.subplots(figsize=(6,2.7));ax.plot(stab.comparison,stab.ari,'o-',color=PAL[1]);ax.axhline(stab.ari.mean(),color=PAL[0],ls='--',label=f'Mean ARI {stab.ari.mean():.3f}');ax.set_ylim(-.05,1.05);ax.set_xlabel('Paired 8,000-household subsample');ax.set_ylabel('Adjusted Rand index');ax.legend(frameon=False);save('10_stability.png')
    # 12 Ward vs K-means, identical comparison subset
    ward=pd.read_csv(out/'ward_comparison_same_sample.csv');fig,ax=plt.subplots(figsize=(7.3,3.1))
    for i,(name,sub) in enumerate(ward.groupby('algorithm')):
        sub=sub.sort_values('k');ax.plot(sub.k,sub.silhouette,'o-',label=name,color=PAL[i],lw=2)
    ax.set_xlabel('Cluster count (k)');ax.set_ylabel('Silhouette score (same 2,100 households)');ax.set_xticks(range(2,7));ax.legend(frameon=False,fontsize=8);save('12_ward_benchmark.png')
    # 11 per-cluster sample percent
    fig,ax=plt.subplots(figsize=(6,2.7));ax.bar([f'Segment {c}' for c in profile.cluster],profile.sample_pct,color=PAL[:len(profile)]);
    for i,v in enumerate(profile.sample_pct):ax.text(i,v+.4,f'{v:.1f}%',ha='center',fontsize=9)
    ax.set_ylabel('Sample share (%)');ax.set_ylim(0,profile.sample_pct.max()*1.18);save('11_segment_sizes.png')


def main():
    here=Path(__file__).resolve().parents[1]
    p=argparse.ArgumentParser()
    p.add_argument('--data',type=Path,default=here/'data'/'CSV_data_HH_Cons_exp_22_23.zip')
    p.add_argument('--out',type=Path,default=here/'results')
    p.add_argument('--fit-limit',type=int,default=45000)
    args=p.parse_args()
    if not args.data.is_file(): raise SystemExit('Original dataset ZIP not found; place inside data/ or use --data PATH')
    args.out.mkdir(parents=True,exist_ok=True)
    raw,meta=extract(args.data)
    raw,selected,comparison,profile,sector_tab,desc,report,visual_data=model_fit(raw,args.out,meta,max_n=args.fit_limit)
    figures(raw,selected,comparison,profile,sector_tab,desc,report,visual_data,args.out)
    print('DONE:',report['best_algorithm'],'k=',report['best_k'],'silhouette=',report['best_silhouette'])
    print('Outputs:',args.out)

if __name__=='__main__':main()
