#include <bits/stdc++.h>
using namespace std;
long long merge_count_reverse(vector<int>& a, int l, int r){
    if(r-l<=1) return 0;
    int m=(l+r)/2;
    long long cnt=merge_count_reverse(a,l,m)+merge_count_reverse(a,m,r);
    int j=m;
    for(int i=l;i<m;++i){
        while(j<r && (long long)a[i] > 2LL*a[j]) j++;
        cnt += (j-m);
    }
    vector<int> tmp; int i=l; j=m;
    while(i<m && j<r){ if(a[i]<=a[j]) tmp.push_back(a[i++]); else tmp.push_back(a[j++]); }
    while(i<m) tmp.push_back(a[i++]);
    while(j<r) tmp.push_back(a[j++]);
    for(int k=l;k<r;k++) a[k]=tmp[k-l];
    return cnt;
}
long long brute_reverse_pairs(vector<int> a){
    long long cnt=0;
    for(int i=0;i<a.size();++i) for(int j=i+1;j<a.size();++j) if((long long)a[i] > 2LL*a[j]) cnt++;
    return cnt;
}
long long optimal_reverse_pairs(vector<int> a){
    return merge_count_reverse(a,0,a.size());
}
int main(){return 0;}