#include <bits/stdc++.h>
using namespace std;
pair<int,int> brute_find_repeating_missing(vector<int>& a){
    int n=a.size();
    vector<int> cnt(n+1,0);
    for(int x:a) cnt[x]++;
    int rep=-1, miss=-1;
    for(int i=1;i<=n;i++){ if(cnt[i]==2) rep=i; if(cnt[i]==0) miss=i; }
    return {rep, miss};
}
pair<int,int> optimal_find_repeating_missing(vector<int>& a){
    int n=a.size();
    long long sum=0, sumsq=0;
    for(int i=1;i<=n;i++){ sum += i; sumsq += 1LL*i*i; }
    long long s1=0, s2=0;
    for(int x:a){ s1 += x; s2 += 1LL*x*x; }
    long long diff = s1 - sum; // A - B
    long long diff2 = s2 - sumsq; // A^2 - B^2 = (A-B)(A+B)
    long long sumAB = diff2 / diff;
    long long A = (diff + sumAB)/2;
    long long B = A - diff;
    return { (int)A, (int)B };
}
int main(){return 0;}